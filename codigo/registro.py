"""Registro en JSON de los e-cheques ya utilizados, para no repetirlos en nuevos pagos."""
import json
import os
import sys
import tempfile
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import List, Optional, Tuple

from core import Check, DataError

FILE_NAME = "registro_echeques.json"
VERSION = 1


def default_path() -> Path:
    """Junto al programa: es visible, se puede copiar y se mueve con él."""
    base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    return base / FILE_NAME


def check_key(check: Check) -> tuple:
    """Identidad de un e-cheque entre archivos distintos.

    No incluye la fila (cambia entre listados) ni el recibo (puede completarse
    después). Filas con la misma clave son intercambiables y se cuentan.
    """
    return (check.reference, check.client_id, check.amount, check.date_f, check.date_g)


def _text(cents: int) -> str:
    return "{}.{:02d}".format(cents // 100, cents % 100)


def _cents(text) -> int:
    try:
        number = Decimal(str(text)) * 100
        if not number.is_finite() or number < 0 or number != number.to_integral_value():
            raise ValueError
        return int(number)
    except (InvalidOperation, ValueError) as exc:
        raise ValueError("importe inválido: {!r}".format(text)) from exc


@dataclass
class RegisteredPayment:
    id: int
    registered: str
    target: int
    client_id: Optional[str]
    file: str
    sheet: str
    checks: List[Check]
    voided: Optional[str] = None
    operation: int = 0
    name: str = ""

    @property
    def total(self) -> int:
        return sum(check.amount for check in self.checks)

    @property
    def transfer(self) -> int:
        return self.target - self.total


@dataclass
class Operation:
    """Una búsqueda guardada: los pagos registrados juntos."""
    id: int
    name: str
    payments: List[RegisteredPayment]

    @property
    def label(self) -> str:
        return self.name or "Operación {}".format(self.id)

    @property
    def registered(self) -> str:
        return self.payments[0].registered

    @property
    def file(self) -> str:
        return self.payments[0].file

    @property
    def sheet(self) -> str:
        return self.payments[0].sheet

    @property
    def checks(self) -> List[Check]:
        return [check for payment in self.payments for check in payment.checks]

    @property
    def target(self) -> int:
        return sum(payment.target for payment in self.payments)

    @property
    def total(self) -> int:
        return sum(payment.total for payment in self.payments)

    @property
    def status(self) -> str:
        voided = sum(1 for payment in self.payments if payment.voided)
        if not voided:
            return "Registrada"
        return "Anulada" if voided == len(self.payments) else "Parcialmente anulada"


def _check_to_json(check: Check) -> dict:
    return {"fila": check.row, "referencia": check.reference, "fecha_f": check.date_f,
            "fecha_g": check.date_g, "id_cliente": check.client_id, "cliente": check.client_name,
            "importe": _text(check.amount), "recibo": check.receipt}


def _check_from_json(item: dict) -> Check:
    amount = _cents(item["importe"])
    if amount <= 0:
        raise ValueError("importe inválido: {!r}".format(item["importe"]))
    return Check(int(item["fila"]), str(item["referencia"]), str(item["fecha_f"]), str(item["fecha_g"]),
                 str(item["id_cliente"]), str(item["cliente"]), amount, str(item["recibo"]))


def _payment_to_json(payment: RegisteredPayment) -> dict:
    data = {"id": payment.id, "operacion": payment.operation, "nombre_operacion": payment.name,
            "registrado": payment.registered, "cliente_filtro": payment.client_id,
            "archivo": payment.file, "hoja": payment.sheet,
            "importe_solicitado": _text(payment.target), "total_echeques": _text(payment.total),
            "transferencia": _text(payment.transfer)}
    if payment.voided:
        data["anulado"] = payment.voided
    data["echeques"] = [_check_to_json(check) for check in payment.checks]
    return data


def _payment_from_json(item: dict) -> RegisteredPayment:
    client = item["cliente_filtro"]
    return RegisteredPayment(int(item["id"]), str(item["registrado"]), _cents(item["importe_solicitado"]),
                             None if client is None else str(client), str(item["archivo"]), str(item["hoja"]),
                             [_check_from_json(check) for check in item["echeques"]],
                             item.get("anulado") or None,
                             int(item.get("operacion", item["id"])), str(item.get("nombre_operacion", "")))


def format_timestamp(value: str) -> str:
    try:
        return datetime.fromisoformat(value).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return value


class Registry:
    """Pagos registrados. Un archivo ilegible nunca se toma por vacío: eso
    permitiría reutilizar e-cheques sin aviso."""

    def __init__(self, path, payments: Optional[List[RegisteredPayment]] = None):
        self.path = Path(path)
        self.payments = payments if payments is not None else []

    @classmethod
    def load(cls, path) -> "Registry":
        path = Path(path)
        if not path.exists():
            return cls(path)
        def problem(detail):
            return DataError("No se pudo leer el registro de e-cheques utilizados ({}). {} "
                             "No se modificó el archivo: corregilo o movelo antes de continuar.".format(path, detail))
        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
        except (OSError, ValueError) as exc:
            raise problem(str(exc)) from exc
        try:
            version = data["version"]
            payments = [_payment_from_json(item) for item in data["pagos"]] if version == VERSION else []
        except (KeyError, TypeError, ValueError, AttributeError) as exc:
            raise problem("Formato inesperado: {}.".format(exc)) from exc
        if version != VERSION:
            raise problem("Fue creado con otra versión del programa.")
        if len({payment.id for payment in payments}) != len(payments):
            raise problem("Hay pagos con el mismo número.")
        return cls(path, payments)

    def save(self) -> None:
        data = {"version": VERSION, "pagos": [_payment_to_json(payment) for payment in self.payments]}
        self.path.parent.mkdir(parents=True, exist_ok=True)
        handle, temp = tempfile.mkstemp(dir=str(self.path.parent), prefix=self.path.name + ".", suffix=".tmp")
        try:
            with os.fdopen(handle, "w", encoding="utf-8") as out:
                json.dump(data, out, ensure_ascii=False, indent=2)
                out.write("\n")
                out.flush()
                os.fsync(out.fileno())
            os.replace(temp, self.path)
        except BaseException:
            try:
                os.unlink(temp)
            except OSError:
                pass
            raise

    def used(self) -> Counter:
        """Cuántas veces se usó cada clave, sin contar pagos anulados."""
        return Counter(check_key(check) for payment in self.payments if not payment.voided
                       for check in payment.checks)

    def split(self, checks: List[Check]) -> Tuple[List[Check], List[Check]]:
        """(disponibles, ya utilizados). Con claves repetidas se descuentan tantas filas como usos."""
        pending = self.used()
        available, used = [], []
        for check in checks:
            key = check_key(check)
            if pending[key] > 0:
                pending[key] -= 1
                used.append(check)
            else:
                available.append(check)
        return available, used

    def conflicts(self, checks: List[Check], selected: List[Check]) -> List[Check]:
        """Elegidos que ya no están disponibles (p. ej. registrados desde otra ventana)."""
        free = Counter(check_key(check) for check in self.split(checks)[0])
        missing = []
        for check in selected:
            key = check_key(check)
            if free[key] > 0:
                free[key] -= 1
            else:
                missing.append(check)
        return missing

    def operations(self) -> List[Operation]:
        groups = {}
        for payment in self.payments:
            groups.setdefault(payment.operation, []).append(payment)
        return [Operation(number, payments[0].name, payments) for number, payments in groups.items()]

    def add(self, outcomes, file: str, sheet: str, when: Optional[datetime] = None,
            name: str = "") -> List[RegisteredPayment]:
        """Agrega, como una operación, los pagos con e-cheques; los demás no tienen nada que registrar."""
        stamp = (when or datetime.now()).isoformat(timespec="seconds")
        number = max((payment.id for payment in self.payments), default=0)
        operation = max((payment.operation for payment in self.payments), default=0) + 1
        name = " ".join(str(name).split())[:100]
        added = []
        for outcome in outcomes:
            result = outcome.result
            if result is None or not result.checks:
                continue
            number += 1
            added.append(RegisteredPayment(number, stamp, result.target, outcome.payment.client_id,
                                           file, sheet, list(result.checks), None, operation, name))
        rows = [check.row for payment in added for check in payment.checks]
        if len(set(rows)) != len(rows):
            raise RuntimeError("Un e-cheque quedó asignado a más de un pago.")
        self.payments.extend(added)
        return added

    def void_operation(self, operation_id: int, when: Optional[datetime] = None) -> List[RegisteredPayment]:
        """Anula todos los pagos activos de una operación."""
        members = [payment for payment in self.payments if payment.operation == operation_id]
        if not members:
            raise DataError("No existe la operación {} en el registro.".format(operation_id))
        active = [payment for payment in members if not payment.voided]
        if not active:
            raise DataError("La operación {} ya estaba anulada.".format(operation_id))
        stamp = (when or datetime.now()).isoformat(timespec="seconds")
        for payment in active:
            payment.voided = stamp
        return active

    def void(self, payment_id: int, when: Optional[datetime] = None) -> RegisteredPayment:
        """Anula un pago: sus e-cheques vuelven a estar disponibles. El historial se conserva."""
        payment = next((item for item in self.payments if item.id == payment_id), None)
        if payment is None:
            raise DataError("No existe el pago {} en el registro.".format(payment_id))
        if payment.voided:
            raise DataError("El pago {} ya estaba anulado.".format(payment_id))
        payment.voided = (when or datetime.now()).isoformat(timespec="seconds")
        return payment
