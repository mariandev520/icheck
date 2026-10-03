"""Lectura de Excel y selección de e-cheques, sin dependencias externas."""
import math
import heapq
from bisect import bisect_left
import posixpath
import re
import time
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Callable, List, Optional, Tuple
from xml.etree import ElementTree as ET


SPLIT_SECONDS = 10.0  # tiempo máximo para repartir e-cheques entre varios pagos


class DataError(ValueError):
    pass


def parse_amount(text: str) -> int:
    """Entrada humana: formato argentino o decimal con punto, en centavos."""
    value = str(text).strip().replace("\u00a0", "").replace(" ", "")
    value = re.sub(r"^(?:ARS|\$)", "", value, flags=re.I)
    if re.fullmatch(r"\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?", value):
        normalized = value.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d+(?:,\d{1,2})?", value):
        normalized = value.replace(",", ".")
    elif re.fullmatch(r"\d+\.\d{1,2}", value):
        normalized = value
    else:
        raise DataError("Ingresá un importe como 500.000 o 500.000,50 (hasta dos decimales).")
    result = int(Decimal(normalized) * 100)
    if result <= 0:
        raise DataError("El importe debe ser mayor que cero.")
    if result > 10**14:
        raise DataError("El importe supera el máximo admitido: 1.000.000.000.000,00.")
    return result


def numeric_cents(value: str) -> int:
    """Número XML: tolera el ruido binario de Excel, no redondea milésimas reales."""
    try:
        number = Decimal(value)
        if not number.is_finite() or number <= 0 or number > Decimal(10**12):
            raise DataError("importe fuera de rango o no positivo")
        rounded = number.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if abs(number - rounded) > Decimal("0.00001"):
            raise DataError("importe con más de dos decimales")
        cents = int(rounded * 100)
        if cents <= 0:
            raise DataError("importe no positivo")
        return cents
    except (InvalidOperation, ValueError) as exc:
        if isinstance(exc, DataError):
            raise
        raise DataError("importe no numérico") from exc


def money(cents: int) -> str:
    whole, fraction = divmod(abs(cents), 100)
    sign = "-" if cents < 0 else ""
    return "{}$ {},{:02d}".format(sign, format(whole, ",").replace(",", "."), fraction)


@dataclass(frozen=True)
class Check:
    row: int
    reference: str
    date_f: str
    date_g: str
    client_id: str
    client_name: str
    amount: int
    receipt: str


@dataclass
class ImportResult:
    checks: List[Check]
    notes: List[str]
    rejected: List[str]
    sheet: str


@dataclass
class SearchResult:
    checks: List[Check]
    target: int
    optimal: bool
    elapsed: float
    strategy: str = "exact"
    completed: bool = True
    date_field: str = "G"

    @property
    def total(self) -> int:
        return sum(check.amount for check in self.checks)

    @property
    def transfer(self) -> int:
        return self.target - self.total


@dataclass(frozen=True)
class Payment:
    """Un pago a cubrir, opcionalmente limitado a los e-cheques de un cliente."""
    target: int
    client_id: Optional[str] = None
    client_name: str = ""

    @property
    def label(self) -> str:
        if self.client_id is None:
            return "Todos los clientes"
        return " · ".join(part for part in (self.client_id, self.client_name) if part)


@dataclass
class PaymentOutcome:
    payment: Payment
    result: Optional[SearchResult]  # None: no se llegó a calcular (búsqueda detenida)


def parse_date(value: str) -> Optional[date]:
    """Fecha humana opcional; no corrige silenciosamente fechas imposibles."""
    value = str(value).strip()
    if not value:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    raise DataError("Ingresá una fecha válida como 25/09/2026 o dejá el campo vacío.")


def check_date(check: Check, date_field: str = "G") -> Optional[date]:
    if date_field not in ("F", "G"):
        raise DataError("La fecha debe corresponder a la columna F o G.")
    try:
        return parse_date(check.date_f if date_field == "F" else check.date_g)
    except DataError:
        return None


def week_start(value: Optional[date]) -> Optional[date]:
    return value - timedelta(days=value.weekday()) if value else None


def batch_label(check: Check, date_field: str = "G") -> str:
    start = week_start(check_date(check, date_field))
    return "{:%d/%m/%Y}–{:%d/%m/%Y}".format(start, start + timedelta(days=6)) if start else "Sin fecha válida"


def filter_checks(checks: List[Check], date_field: str = "G",
                  start: Optional[date] = None, end: Optional[date] = None):
    """Rango inclusivo. Sin rango conserva incluso registros con fecha inválida."""
    if date_field not in ("F", "G"):
        raise DataError("La fecha debe corresponder a la columna F o G.")
    if start and end and start > end:
        raise DataError("La fecha Desde no puede ser posterior a Hasta.")
    selected, excluded_dates = [], 0
    for check in checks:
        value = check_date(check, date_field)
        if (start or end) and value is None:
            excluded_dates += 1
        elif not (start and value < start or end and value > end):
            selected.append(check)
    return selected, excluded_dates


def find_progressive(checks: List[Check], target: int, cancel=None,
                     progress: Optional[Callable[[str], None]] = None,
                     date_field: str = "G", prefer_previous: bool = True) -> SearchResult:
    """Selección descendente, sin exceder el saldo ni reutilizar filas.

    Exacto al saldo primero. En su defecto, favorece la semana anterior
    disponible al último elegido; dentro de ella toma el mayor que entra.
    Sin semana anterior utilizable vuelve al mayor disponible. La restricción
    de importe descendente rige también al priorizar fechas. No es un óptimo
    global de suma o cantidad: un saldo pendiente no se presenta como tal.
    """
    if not isinstance(target, int) or target <= 0:
        raise DataError("El importe solicitado debe ser un número positivo de centavos.")
    if date_field not in ("F", "G"):
        raise DataError("La fecha debe corresponder a la columna F o G.")
    if any(not isinstance(c.amount, int) or c.amount <= 0 for c in checks):
        raise DataError("Todos los e-cheques deben tener importes positivos en centavos.")
    started = time.monotonic()
    candidates = [c for c in checks if c.amount <= target]
    dates = [check_date(c, date_field) for c in candidates]
    batches = [week_start(d) for d in dates]
    # Índice de fila y posición original resuelven empates sin mezclar identidades.
    entries = [(-c.amount, -d.toordinal() if d else 0, c.row, i)
               for i, (c, d) in enumerate(zip(candidates, dates))]
    available = list(entries)
    heapq.heapify(available)
    by_week = {}
    for entry, batch in zip(entries, batches):
        if batch is not None:
            by_week.setdefault(batch, []).append(entry)
    for heap in by_week.values():
        heapq.heapify(heap)
    weeks = sorted(by_week)
    used, selected = set(), []
    remaining, ceiling, previous = target, target, None
    completed, last_report = True, started

    def first(heap, cap):
        while heap and (heap[0][3] in used or -heap[0][0] > cap):
            heapq.heappop(heap)
        return heap[0] if heap else None

    while remaining:
        if cancel is not None and cancel.is_set():
            completed = False
            break
        cap = min(remaining, ceiling)
        best = first(available, cap)
        if best is None:
            break
        older = None
        if prefer_previous and previous is not None:
            position = bisect_left(weeks, previous) - 1
            while position >= 0:
                older = first(by_week[weeks[position]], cap)
                if older is not None:
                    break
                weeks.pop(position)
                position -= 1
        # Exacto manda; para exactos empatados se prefiere la tanda anterior.
        if older is not None and (-best[0] != remaining or -older[0] == remaining):
            best = older
        index = best[3]
        used.add(index)
        chosen = candidates[index]
        selected.append(chosen)
        remaining -= chosen.amount
        ceiling = chosen.amount
        previous = batches[index]
        if progress and time.monotonic() - last_report >= 0.2:
            progress("{} e-cheques elegidos · saldo por cubrir: {}".format(len(selected), money(remaining)))
            last_report = time.monotonic()
    total = target - remaining
    proven = remaining == 0 or (completed and total == sum(c.amount for c in candidates))
    return SearchResult(selected, target, proven, time.monotonic() - started,
                        "progressive", completed, date_field)


def _xml(archive: zipfile.ZipFile, name: str):
    info = archive.getinfo(name)
    if info.file_size > 60 * 1024 * 1024:
        raise DataError("La hoja es demasiado grande para esta versión (máximo 60 MB de XML).")
    root = ET.fromstring(archive.read(name))
    for element in root.iter():
        element.tag = element.tag.rsplit("}", 1)[-1]
    return root


def _part(base: str, target: str) -> str:
    result = posixpath.normpath(posixpath.join(base, target)) if not target.startswith("/") else target[1:]
    if result.startswith("../"):
        raise DataError("El archivo contiene una referencia de hoja inválida.")
    return result


def _relationships(archive):
    rels = _xml(archive, "xl/_rels/workbook.xml.rels")
    return {rel.get("Id"): (_part("xl", rel.get("Target", "")), rel.get("Type", ""))
            for rel in rels if rel.get("TargetMode") != "External"}


def _sheets(archive):
    workbook = _xml(archive, "xl/workbook.xml")
    rels = _relationships(archive)
    result = []
    for sheet in workbook.findall("sheets/sheet"):
        rel_id = next((value for key, value in sheet.attrib.items() if key.endswith("}id")), None)
        if rel_id in rels and rels[rel_id][1].endswith("/worksheet"):
            result.append((sheet.get("name", "Hoja"), rels[rel_id][0]))
    props = workbook.find("workbookPr")
    date1904 = props is not None and props.get("date1904") in ("1", "true")
    return result, date1904, rels


def list_sheets(path) -> List[str]:
    try:
        with zipfile.ZipFile(path) as archive:
            sheets, _, _ = _sheets(archive)
            if not sheets:
                raise DataError("El archivo no contiene hojas de cálculo.")
            return [name for name, _ in sheets]
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        raise DataError("No se pudo leer el Excel. Usá un archivo .xlsx sin contraseña.") from exc


def _date(value: str, numeric: bool, date1904: bool) -> Tuple[str, bool]:
    if not value:
        return "", True
    if numeric:
        try:
            days = float(value)
            base = datetime(1904, 1, 1) if date1904 else datetime(1899, 12, 30)
            return (base + timedelta(days=days)).strftime("%d/%m/%Y"), False
        except (ValueError, OverflowError):
            return value, True
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%d-%m-%Y"):
        try:
            return datetime.strptime(value, fmt).strftime("%d/%m/%Y"), False
        except ValueError:
            pass
    return value, True


def read_excel(path, sheet_name: Optional[str] = None) -> ImportResult:
    """Lee F/G/J/R/V y referencias C/W; conserva fechas marcadas e IDs textuales."""
    try:
        with zipfile.ZipFile(path) as archive:
            sheets, date1904, rels = _sheets(archive)
            if not sheets:
                raise DataError("No hay hojas de cálculo en este archivo.")
            selected = next((item for item in sheets if item[0] == sheet_name), None) if sheet_name else sheets[0]
            if selected is None:
                raise DataError("No se encontró la hoja seleccionada.")
            shared = []
            shared_path = next((part for part, kind in rels.values() if kind.endswith("/sharedStrings")), None)
            if shared_path:
                shared = ["".join(t.text or "" for t in item.iter("t")) for item in _xml(archive, shared_path)]
            # Algunos exportadores omiten la relación de sharedStrings.
            elif "xl/sharedStrings.xml" in archive.namelist():
                shared = ["".join(t.text or "" for t in item.iter("t"))
                          for item in _xml(archive, "xl/sharedStrings.xml")]
            root = _xml(archive, selected[1])
            checks, notes, rejected = [], [], []
            for row in root.findall("sheetData/row"):
                cells, types, formulas = {}, {}, set()
                row_number = int(row.get("r", "0"))
                for cell in row.findall("c"):
                    col = re.sub(r"\d", "", cell.get("r", ""))
                    kind = cell.get("t", "n")
                    value = cell.findtext("v", "")
                    if kind == "s":
                        value = shared[int(value)]
                    elif kind == "inlineStr":
                        value = "".join(t.text or "" for t in cell.iter("t"))
                    cells[col], types[col] = value.strip(), kind
                    if cell.find("f") is not None:
                        formulas.add(col)
                if not any(cells.get(col) for col in ("C", "F", "G", "J", "R", "V", "W")):
                    continue
                raw_amount = cells.get("V", "")
                if raw_amount.lower() in ("importe", "monto", "valor", "total", "amount"):
                    continue
                if cells.get("J", "").lower() in ("cliente", "id cliente", "id de cliente", "código cliente"):
                    continue
                try:
                    if not raw_amount:
                        raise DataError("falta el importe en V")
                    amount = numeric_cents(raw_amount) if types.get("V") == "n" else parse_amount(raw_amount)
                    client_id = cells.get("J", "")
                    if not client_id or not client_id.isascii() or not client_id.isdigit():
                        raise DataError("falta un ID numérico de cliente en J")
                    if types.get("J") == "n":
                        client_id = client_id.zfill(5)
                    if len(client_id) != 5:
                        notes.append("Fila {}: ID {} conservado tal como figura en Excel.".format(row_number, client_id))
                    date_f, marked_f = _date(cells.get("F", ""), types.get("F") == "n", date1904)
                    date_g, marked_g = _date(cells.get("G", ""), types.get("G") == "n", date1904)
                    for col, date, marked in (("F", date_f, marked_f), ("G", date_g, marked_g)):
                        if marked:
                            notes.append("Fila {}, fecha {}: {}. Registro incluido sin cambiar la fecha.".format(
                                row_number, col, date or "vacía"))
                    if "V" in formulas:
                        notes.append("Fila {}: importe tomado del resultado guardado de una fórmula; no se recalcula.".format(row_number))
                    checks.append(Check(row_number, cells.get("C", ""), date_f, date_g, client_id,
                                        cells.get("R", ""), amount, cells.get("W", "")))
                    if len(checks) > 20000:
                        raise DataError("Esta versión admite hasta 20.000 registros por hoja.")
                except DataError as exc:
                    if len(checks) > 20000:
                        raise
                    rejected.append("Fila {}: {}.".format(row_number, exc))
            return ImportResult(checks, notes, rejected, selected[0])
    except (zipfile.BadZipFile, KeyError, ET.ParseError, IndexError) as exc:
        raise DataError("El Excel está dañado o tiene un formato no compatible. Guardalo como .xlsx sin contraseña.") from exc


def _greedy(values, target, order):
    total, picked = 0, 0
    for i in order:
        if total + values[i] <= target:
            total += values[i]
            picked |= 1 << i
    return total, picked


def find_combination(checks: List[Check], target: int, cancel=None,
                     progress: Optional[Callable[[str], None]] = None,
                     memory_mb: int = 160) -> SearchResult:
    """Máxima suma <= objetivo. Cancelar devuelve una suma válida, marcada no óptima.

    DP exacta con bitsets y puntos de reconstrucción de memoria acotada.
    Para objetivos que exceden ese presupuesto se usa ramificación exacta.
    No hay límite de tiempo ni aproximación silenciosa.
    """
    if not isinstance(target, int) or target <= 0:
        raise DataError("El importe solicitado debe ser un número positivo de centavos.")
    if any(not isinstance(check.amount, int) or check.amount <= 0 for check in checks):
        raise DataError("Todos los e-cheques deben tener importes positivos en centavos.")
    started = time.monotonic()
    candidates = sorted((check for check in checks if check.amount <= target), key=lambda item: item.amount)
    def stopped():
        return cancel is not None and cancel.is_set()
    def report(message):
        if progress:
            progress(message)
    def finish(indices, optimal):
        selected = sorted((candidates[i] for i in indices), key=lambda item: item.row)
        result = SearchResult(selected, target, optimal, time.monotonic() - started, completed=optimal)
        if result.total > target:
            raise RuntimeError("La selección superó el importe solicitado.")
        return result
    if not candidates:
        return finish([], True)
    total = sum(check.amount for check in candidates)
    if total <= target:
        return finish(range(len(candidates)), True)
    for i, check in enumerate(candidates):
        if check.amount == target:
            return finish([i], True)
    divisor = 0
    for check in candidates:
        divisor = math.gcd(divisor, check.amount)
    values = [check.amount // divisor for check in candidates]
    limit = target // divisor
    count = len(values)
    seed_total, seed_picked = max(
        (_greedy(values, limit, range(count)), _greedy(values, limit, range(count - 1, -1, -1))),
        key=lambda item: item[0])
    def from_mask(mask, optimal):
        return finish([i for i in range(count) if mask & (1 << i)], optimal)
    if seed_total == limit:
        return from_mask(seed_picked, True)
    if stopped():
        return from_mask(seed_picked, False)
    block = max(1, math.isqrt(count))
    # CPython: ~30 bits por 4 bytes, más checkpoints, bloque y temporarios.
    estimated = (limit + 1) / 7.5 * (math.ceil(count / block) + block + 8)
    if estimated <= memory_mb * 1024 * 1024:
        mask = (1 << (limit + 1)) - 1
        reachable, checkpoints, processed = 1, [], 0
        last_report = started
        for i, value in enumerate(values):
            if stopped():
                break
            if i % block == 0:
                checkpoints.append(reachable)
            reachable |= (reachable << value) & mask
            processed = i + 1
            now = time.monotonic()
            if now - last_report >= 0.2:
                report("Buscando la combinación más cercana… {} de {} e-cheques analizados.".format(processed, count))
                last_report = now
            if reachable >> limit & 1:
                break
        best = reachable.bit_length() - 1
        optimal = processed == count or best == limit
        if seed_total > best:
            return from_mask(seed_picked, optimal)
        remaining, picked = best, []
        for batch in range((processed - 1) // block, -1, -1):
            begin, end = batch * block, min((batch + 1) * block, processed)
            states = [checkpoints[batch]]
            for i in range(begin, end - 1):
                states.append(states[-1] | ((states[-1] << values[i]) & mask))
            for i in range(end - 1, begin - 1, -1):
                if not (states[i - begin] >> remaining & 1):
                    picked.append(i)
                    remaining -= values[i]
            if remaining == 0:
                break
        if remaining:
            raise RuntimeError("No se pudo reconstruir la combinación.")
        return finish(picked, optimal)

    report("Analizando combinaciones. Podés detener la búsqueda y ver lo encontrado.")
    # Orden descendente, manteniendo el índice original de cada registro.
    order = sorted(range(count), key=lambda i: values[i], reverse=True)
    suffix, suffix_mask = [0] * (count + 1), [0] * (count + 1)
    for i in range(count - 1, -1, -1):
        suffix[i] = suffix[i + 1] + values[order[i]]
        suffix_mask[i] = suffix_mask[i + 1] | (1 << order[i])
    best, best_mask = seed_total, seed_picked
    pending, seen = [(0, 0, 0)], set()
    visited, last_report = 0, started
    while pending:
        if stopped():
            return from_mask(best_mask, False)
        i, subtotal, selected = pending.pop()
        if subtotal > best:
            best, best_mask = subtotal, selected
            if best == limit:
                return from_mask(best_mask, True)
        if i == count or subtotal + suffix[i] <= best:
            continue
        if subtotal + suffix[i] <= limit:
            best, best_mask = subtotal + suffix[i], selected | suffix_mask[i]
            if best == limit:
                return from_mask(best_mask, True)
            continue
        state = (i, subtotal)
        if state in seen:
            continue
        if len(seen) >= 250000:
            seen.clear()
        seen.add(state)
        index = order[i]
        pending.append((i + 1, subtotal, selected))
        if subtotal + values[index] <= limit:
            pending.append((i + 1, subtotal + values[index], selected | (1 << index)))
        visited += 1
        if visited % 2048 == 0 and time.monotonic() - last_report > 0.2:
            report("Buscando… mejor total hasta ahora: {}. Nunca supera lo solicitado.".format(money(best * divisor)))
            last_report = time.monotonic()
    return from_mask(best_mask, True)


def _split_search(items, targets, ceilings, seed, cancel, progress, deadline):
    """Reparto exacto de e-cheques entre pagos por ramificación y poda.

    items: [(importe, pagos admitidos)] en orden descendente. ceilings[j]: lo máximo
    que el pago j puede sumar aun solo (cota). seed: pago asignado a cada item (-1: ninguno).
    Maximiza el total cubierto y, a igual total, la cantidad de pagos exactos.
    Devuelve (asignación, estado): "proven", "cancelled" o "timeout".
    """
    k, n = len(targets), len(items)
    amt = [item[0] for item in items]
    opts = [item[1] for item in items]
    reach = [[0] * (n + 1) for _ in range(k)]
    total_reach = [0] * (n + 1)
    for t in range(n - 1, -1, -1):
        total_reach[t] = total_reach[t + 1] + amt[t]
        for j in range(k):
            reach[j][t] = reach[j][t + 1] + (amt[t] if j in opts[t] else 0)
    can_be_exact = [ceilings[j] == targets[j] for j in range(k)]
    cap = list(ceilings)
    covered = 0
    chosen = [-2] * n
    best = list(seed)
    best_cov = sum(a for a, j in zip(amt, seed) if j >= 0)
    sums = [0] * k
    for a, j in zip(amt, seed):
        if j >= 0:
            sums[j] += a
    best_ex = sum(1 for j in range(k) if can_be_exact[j] and sums[j] == targets[j])
    ceiling_total = sum(ceilings)
    perfect = (ceiling_total, sum(can_be_exact))

    def bound(t):
        room, possible = 0, 0
        for j in range(k):
            r = reach[j][t]
            room += cap[j] if cap[j] < r else r
            if can_be_exact[j] and cap[j] <= r:
                possible += 1
        return covered + min(room, total_reach[t]), possible

    state = {"best_cov": best_cov, "best_ex": best_ex, "done": False}

    def make(t):
        if state["done"]:
            return []
        total, possible = bound(t)
        if total < state["best_cov"] or (total == state["best_cov"] and possible <= state["best_ex"]):
            return []
        if t == n:
            exact = sum(1 for j in range(k) if can_be_exact[j] and cap[j] == 0)
            if (covered, exact) > (state["best_cov"], state["best_ex"]):
                state["best_cov"], state["best_ex"] = covered, exact
                best[:] = chosen
                if (covered, exact) >= perfect:
                    state["done"] = True
            return []
        floor = -1
        if t and amt[t] == amt[t - 1] and opts[t] == opts[t - 1]:
            floor = k if chosen[t - 1] == -1 else chosen[t - 1]  # iguales: se usan en orden
        fits = sorted((j for j in opts[t] if cap[j] >= amt[t] and j >= floor), key=lambda j: cap[j])
        return [-1] + fits[::-1]

    options = [None] * (n + 1)
    options[0] = make(0)
    t, steps, last_report, status = 0, 0, time.monotonic(), "proven"
    while t >= 0 and not state["done"]:
        steps += 1
        if steps % 512 == 0:
            now = time.monotonic()
            if cancel is not None and cancel.is_set():
                status = "cancelled"
                break
            if now > deadline:
                status = "timeout"
                break
            if progress and now - last_report >= 0.2:
                progress("Repartiendo entre los pagos… faltante total hasta ahora: {}.".format(
                    money(sum(targets) - state["best_cov"])))
                last_report = now
        if not options[t]:
            t -= 1
            if t >= 0 and chosen[t] >= 0:
                cap[chosen[t]] += amt[t]
                covered -= amt[t]
            continue
        j = options[t].pop()
        chosen[t] = j
        if j >= 0:
            cap[j] -= amt[t]
            covered += amt[t]
        t += 1
        options[t] = make(t)
    return best, status


def find_payments(checks: List[Check], payments: List[Payment], progressive: bool = True,
                  date_field: str = "G", start: Optional[date] = None, end: Optional[date] = None,
                  prefer_previous: bool = True, cancel=None,
                  progress: Optional[Callable[[str], None]] = None,
                  split_seconds: float = SPLIT_SECONDS) -> List[PaymentOutcome]:
    """Calcula los pagos en el orden dado; un e-cheque elegido no se ofrece a los siguientes.

    Cada pago aplica su cliente y el rango de fechas sobre lo que quedó libre.
    Si se cancela, el pago en curso conserva una selección válida y los
    siguientes quedan sin calcular (result None).

    Con "mejor suma" y varios pagos, además, se busca el reparto entre todos que deje
    el menor faltante total y complete más pagos exactos (split_seconds limita esa
    búsqueda; si se agota, se informa que no se demostró que sea el mejor).
    """
    pool = list(checks)
    outcomes = []
    for number, payment in enumerate(payments, 1):
        if cancel is not None and cancel.is_set():
            outcomes.append(PaymentOutcome(payment, None))
            continue
        records = [c for c in pool if payment.client_id is None or c.client_id == payment.client_id]
        candidates, _ = filter_checks(records, date_field, start, end)
        report = progress
        if progress and len(payments) > 1:
            report = lambda message, n=number: progress("Pago {} de {} · {}".format(n, len(payments), message))
        if progressive:
            result = find_progressive(candidates, payment.target, cancel, report, date_field, prefer_previous)
        else:
            result = find_combination(candidates, payment.target, cancel, report)
            result.date_field = date_field
        taken = {check.row for check in result.checks}
        pool = [c for c in pool if c.row not in taken]
        outcomes.append(PaymentOutcome(payment, result))
    if (progressive or len(payments) < 2 or any(o.result is None or not o.result.completed for o in outcomes)
            or all(o.result.transfer == 0 for o in outcomes)):
        return outcomes
    return _split_between_payments(checks, payments, outcomes, date_field, start, end,
                                   cancel, progress, split_seconds)


class _Budget:
    """Evento de cancelación que además vence al agotarse el tiempo del reparto."""

    def __init__(self, cancel, deadline):
        self.cancel, self.deadline = cancel, deadline

    def is_set(self):
        return (self.cancel is not None and self.cancel.is_set()) or time.monotonic() > self.deadline


def _split_between_payments(checks, payments, outcomes, date_field, start, end, cancel, progress, split_seconds):
    started = time.monotonic()
    budget = _Budget(cancel, started + split_seconds)
    eligible = []
    for payment in payments:
        records = [c for c in checks if payment.client_id is None or c.client_id == payment.client_id]
        eligible.append(filter_checks(records, date_field, start, end)[0])
    # Cota: nadie puede sumar más de lo que lograría solo con todos los e-cheques de su cliente.
    ceilings = []
    for j, payment in enumerate(payments):
        if j == 0:  # el primer pago ya se calculó solo, sobre todos los e-cheques
            ceilings.append(outcomes[0].result.total)
            continue
        solo = find_combination(eligible[j], payment.target, budget,
                                (lambda m, n=j + 1: progress("Cota del pago {} · {}".format(n, m))) if progress else None)
        if not solo.completed:  # sin cotas no se puede demostrar nada: queda el cálculo secuencial
            return _rebuild(checks, payments, outcomes, {c.row: j for j, o in enumerate(outcomes)
                                                         for c in o.result.checks},
                            False, bool(cancel is not None and cancel.is_set()), date_field, started)
        ceilings.append(solo.total)
    baseline = sum(o.result.total for o in outcomes)
    exact_now = sum(1 for o, c, p in zip(outcomes, ceilings, payments) if c == p.target and o.result.total == c)
    unreachable = (baseline == sum(ceilings)
                   and exact_now == sum(1 for c, p in zip(ceilings, payments) if c == p.target))
    proven, stopped = unreachable, False
    assigned = {c.row: j for j, o in enumerate(outcomes) for c in o.result.checks}
    if not unreachable:
        by_row = {c.row: c for group in eligible for c in group}
        order = sorted(by_row.values(), key=lambda c: (-c.amount, c.row))
        rows_by_payment = [{c.row for c in group} for group in eligible]
        items, kept = [], []
        for c in order:
            admitted = tuple(j for j in range(len(payments))
                             if c.row in rows_by_payment[j] and c.amount <= ceilings[j])
            if admitted:
                items.append((c.amount, admitted))
                kept.append(c)
        seed = [assigned.get(c.row, -1) for c in kept]
        best, status = _split_search(items, [p.target for p in payments], ceilings, seed, cancel, progress,
                                     started + split_seconds)
        assigned = {c.row: j for c, j in zip(kept, best) if j >= 0}
        proven, stopped = status == "proven", status == "cancelled"
    return _rebuild(checks, payments, outcomes, assigned, proven, stopped, date_field, started)


def _rebuild(checks, payments, outcomes, assigned, proven, stopped, date_field, started):
    """Resultados por pago a partir de la asignación fila -> pago. El tiempo del reparto va al primero."""
    elapsed = time.monotonic() - started
    rebuilt = []
    for j, (payment, outcome) in enumerate(zip(payments, outcomes)):
        mine = sorted((c for c in checks if assigned.get(c.row) == j), key=lambda c: c.row)
        result = SearchResult(mine, payment.target, proven, outcome.result.elapsed + (elapsed if j == 0 else 0),
                              "exact", not stopped, date_field)
        rebuilt.append(PaymentOutcome(payment, result))
    return rebuilt


def customer_message(result: SearchResult) -> str:
    if not result.checks:
        return ("Para completar el importe de {}, no se seleccionaron e-cheques. "
                "El importe a completar por transferencia es {}.").format(money(result.target), money(result.transfer))
    message = "Para completar el importe de {}, se seleccionaron {} e-cheques por un total de {}. ".format(
        money(result.target), len(result.checks), money(result.total))
    if result.transfer:
        return message + "El importe restante a completar por transferencia es {}.".format(money(result.transfer))
    return message + "El importe queda cubierto; no hace falta completar por transferencia."


def payments_message(outcomes: List[PaymentOutcome]) -> str:
    if len(outcomes) == 1 and outcomes[0].result is not None:
        return customer_message(outcomes[0].result)
    lines = []
    for number, outcome in enumerate(outcomes, 1):
        text = "sin calcular." if outcome.result is None else customer_message(outcome.result)
        lines.append("Pago {} ({}): {}".format(number, outcome.payment.label, text))
    return "\n".join(lines)
