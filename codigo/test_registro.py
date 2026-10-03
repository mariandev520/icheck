import itertools
import json
import random
import tempfile
import threading
import time
import unittest
from dataclasses import replace
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch

from core import (DataError, Payment, PaymentOutcome, SearchResult, find_payments,
                  payments_message)
from core import _split_search
from registro import Registry, check_key, default_path
from test_core import checks_for


def outcome(payment, checks):
    return PaymentOutcome(payment, SearchResult(list(checks), payment.target, True, 0.0))


class RegistryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name) / "registro.json"

    def tearDown(self):
        self.temp.cleanup()

    def test_missing_file_is_an_empty_registry_and_saves_readable_json(self):
        registry = Registry.load(self.path)
        self.assertEqual(registry.payments, [])
        checks = checks_for([30000000, 15000000])
        registry.add([outcome(Payment(50000000, "00123", "Cliente"), checks)], "pagos.xlsx", "Pagos",
                     datetime(2026, 10, 3, 14, 30, 5))
        registry.save()
        data = json.loads(self.path.read_text(encoding="utf-8"))
        payment = data["pagos"][0]
        self.assertEqual(data["version"], 1)
        self.assertEqual((payment["id"], payment["registrado"], payment["cliente_filtro"]),
                         (1, "2026-10-03T14:30:05", "00123"))
        self.assertEqual((payment["importe_solicitado"], payment["total_echeques"], payment["transferencia"]),
                         ("500000.00", "450000.00", "50000.00"))
        self.assertEqual([c["importe"] for c in payment["echeques"]], ["300000.00", "150000.00"])
        again = Registry.load(self.path)
        self.assertEqual(again.payments[0].checks, checks)
        self.assertEqual(again.payments[0].transfer, 5000000)
        self.assertEqual(list(self.path.parent.glob("*.tmp")), [])

    def test_used_checks_are_not_available_even_in_another_listing(self):
        records = checks_for([300, 200, 100])
        registry = Registry(self.path)
        registry.add([outcome(Payment(500), records[:2])], "a.xlsx", "Pagos")
        # Otro Excel: mismas e-cheques en otro orden y con otras filas.
        other = [replace(c, row=c.row + 40) for c in reversed(records)]
        available, used = registry.split(other)
        self.assertEqual([c.amount for c in available], [100])
        self.assertEqual(sorted(c.amount for c in used), [200, 300])

    def test_identical_checks_are_counted_not_all_blocked(self):
        twins = [replace(c, reference="X") for c in checks_for([300, 300, 300])]
        self.assertEqual(len({check_key(c) for c in twins}), 1)
        registry = Registry(self.path)
        registry.add([outcome(Payment(300), twins[:1])], "a.xlsx", "Pagos")
        self.assertEqual(len(registry.split(twins)[0]), 2)
        registry.add([outcome(Payment(300), twins[1:2])], "a.xlsx", "Pagos")
        available, used = registry.split(twins)
        self.assertEqual((len(available), len(used)), (1, 2))

    def test_void_frees_checks_but_keeps_history(self):
        records = checks_for([300, 200])
        registry = Registry(self.path)
        registry.add([outcome(Payment(500), records)], "a.xlsx", "Pagos")
        registry.save()
        registry = Registry.load(self.path)
        self.assertEqual(registry.split(records)[0], [])
        registry.void(1, datetime(2026, 10, 4, 9, 0, 0))
        registry.save()
        registry = Registry.load(self.path)
        self.assertEqual(registry.split(records)[0], records)
        self.assertEqual(registry.payments[0].voided, "2026-10-04T09:00:00")
        self.assertEqual(len(registry.payments), 1)
        with self.assertRaises(DataError):
            registry.void(1)
        with self.assertRaises(DataError):
            registry.void(99)

    def test_ids_continue_after_existing_payments_and_empty_payments_are_skipped(self):
        records = checks_for([300, 200, 100])
        registry = Registry(self.path)
        added = registry.add([outcome(Payment(300), records[:1]), outcome(Payment(5), []),
                              PaymentOutcome(Payment(9), None), outcome(Payment(300), records[1:])],
                             "a.xlsx", "Pagos")
        self.assertEqual([p.id for p in added], [1, 2])
        more = registry.add([outcome(Payment(1), [])], "a.xlsx", "Pagos")
        self.assertEqual(more, [])
        self.assertEqual(registry.add([outcome(Payment(100), [replace(records[0], row=77, amount=100)])],
                                      "a.xlsx", "Pagos")[0].id, 3)

    def test_same_check_in_two_payments_is_refused(self):
        records = checks_for([300])
        registry = Registry(self.path)
        with self.assertRaises(RuntimeError):
            registry.add([outcome(Payment(300), records), outcome(Payment(300), records)], "a.xlsx", "Pagos")
        self.assertEqual(registry.payments, [])

    def test_conflicts_detect_checks_registered_in_the_meantime(self):
        records = checks_for([300, 200, 100])
        registry = Registry(self.path)
        self.assertEqual(registry.conflicts(records, records[:2]), [])
        registry.add([outcome(Payment(300), records[:1])], "a.xlsx", "Pagos")
        self.assertEqual(registry.conflicts(records, records[:2]), records[:1])

    def test_unreadable_registry_is_an_error_never_an_empty_registry(self):
        for content in ("{no es json", "[]", '{"version": 1}', '{"version": 2, "pagos": []}',
                        '{"version": 1, "pagos": [{"id": 1}]}', ""):
            with self.subTest(content=content):
                self.path.write_text(content, encoding="utf-8")
                with self.assertRaises(DataError) as caught:
                    Registry.load(self.path)
                self.assertIn(str(self.path), str(caught.exception))
                self.assertEqual(self.path.read_text(encoding="utf-8"), content)

    def test_duplicate_payment_numbers_and_bad_amounts_are_errors(self):
        registry = Registry(self.path)
        registry.add([outcome(Payment(300), checks_for([300]))], "a.xlsx", "Pagos")
        registry.save()
        data = json.loads(self.path.read_text(encoding="utf-8"))
        broken = json.loads(json.dumps(data))
        broken["pagos"].append(broken["pagos"][0])
        self.path.write_text(json.dumps(broken), encoding="utf-8")
        with self.assertRaises(DataError):
            Registry.load(self.path)
        broken = json.loads(json.dumps(data))
        broken["pagos"][0]["echeques"][0]["importe"] = "0.001"
        self.path.write_text(json.dumps(broken), encoding="utf-8")
        with self.assertRaises(DataError):
            Registry.load(self.path)

    def test_file_edited_with_a_bom_still_loads(self):
        registry = Registry(self.path)
        registry.add([outcome(Payment(300), checks_for([300]))], "a.xlsx", "Pagos")
        registry.save()
        self.path.write_bytes(b"\xef\xbb\xbf" + self.path.read_bytes())
        self.assertEqual(len(Registry.load(self.path).payments), 1)

    def test_default_path_is_next_to_the_program(self):
        self.assertEqual(default_path().name, "registro_echeques.json")
        self.assertEqual(default_path().parent, Path(__file__).resolve().parent)


class PaymentsTests(unittest.TestCase):
    def test_a_check_chosen_for_one_payment_is_not_offered_to_the_next(self):
        records = checks_for([500, 400, 300, 200])
        outcomes = find_payments(records, [Payment(500), Payment(500), Payment(500)])
        chosen = [[c.amount for c in o.result.checks] for o in outcomes]
        self.assertEqual(chosen, [[500], [400], [300, 200]])
        rows = [c.row for o in outcomes for c in o.result.checks]
        self.assertEqual(len(rows), len(set(rows)))
        self.assertEqual([o.result.transfer for o in outcomes], [0, 100, 0])

    def test_one_payment_can_use_several_checks_and_each_has_its_own_client_filter(self):
        records = [replace(c, client_id="A" if i < 2 else "B") for i, c in enumerate(checks_for([300, 200, 400, 100]))]
        outcomes = find_payments(records, [Payment(1000, "A"), Payment(450, "B"), Payment(900)])
        self.assertEqual([c.amount for c in outcomes[0].result.checks], [300, 200])
        self.assertEqual([c.amount for c in outcomes[1].result.checks], [400])
        self.assertEqual([c.amount for c in outcomes[2].result.checks], [100])
        self.assertEqual(outcomes[2].result.transfer, 800)
        self.assertEqual([c.amount for c in find_payments(records, [Payment(100)])[0].result.checks], [100])

    def test_dates_and_both_strategies_apply_to_every_payment(self):
        records = [replace(c, date_g=d) for c, d in zip(checks_for([800, 700, 600, 500]),
                   ['01/09/2026', '08/09/2026', '08/09/2026', '08/09/2026'])]
        window = dict(date_field="G", start=date(2026, 9, 8), end=date(2026, 9, 8))
        exact = find_payments(records, [Payment(1100), Payment(1100)], progressive=False, **window)
        self.assertEqual([c.amount for c in exact[0].result.checks], [600, 500])
        self.assertEqual([c.amount for c in exact[1].result.checks], [700])
        self.assertTrue(all(o.result.date_field == "G" for o in exact))
        progressive = find_payments(records, [Payment(1100)], **window)
        self.assertEqual([c.amount for c in progressive[0].result.checks], [700])

    def test_cancelled_batch_keeps_a_valid_partial_and_marks_the_rest_uncalculated(self):
        event = threading.Event()
        event.set()
        outcomes = find_payments(checks_for([500, 400]), [Payment(500), Payment(400)], cancel=event)
        self.assertEqual([o.result for o in outcomes], [None, None])
        self.assertEqual(find_payments(checks_for([500]), [], cancel=event), [])

    def test_stop_in_the_middle_of_the_batch(self):
        class StopAfter:
            """is_set() pasa a True después de n consultas."""
            def __init__(self, n):
                self.calls, self.n = 0, n

            def is_set(self):
                self.calls += 1
                return self.calls > self.n

        records = checks_for([500, 400, 300])
        outcomes = find_payments(records, [Payment(500), Payment(400), Payment(300)], cancel=StopAfter(3))
        first, second, third = outcomes
        self.assertTrue(first.result.completed)
        self.assertEqual([c.amount for c in first.result.checks], [500])
        self.assertFalse(second.result.completed)
        self.assertIsNone(third.result)
        taken = [c.row for o in outcomes if o.result for c in o.result.checks]
        self.assertEqual(len(taken), len(set(taken)))

    def test_progress_names_the_payment_only_when_there_are_several(self):
        records = checks_for([100] * 50)
        for payments, expected in (([Payment(2000)], False), ([Payment(2000), Payment(2000)], True)):
            seen = []
            with patch("core.time.monotonic", side_effect=itertools.count(0, 1)):
                find_payments(records, payments, progress=seen.append)
            self.assertTrue(seen)
            self.assertEqual(all(m.startswith("Pago ") for m in seen), expected)

    def test_message_for_one_and_for_several_payments(self):
        records = checks_for([30000000, 15000000, 60000000])
        single = find_payments(records, [Payment(50000000)])
        self.assertIn("$ 50.000,00", payments_message(single))
        self.assertFalse(payments_message(single).startswith("Pago"))
        several = find_payments(records, [Payment(50000000), Payment(10000000, "00123", "Cliente")])
        lines = payments_message(several).split("\n")
        self.assertEqual(len(lines), 2)
        self.assertTrue(lines[0].startswith("Pago 1 (Todos los clientes): "))
        self.assertTrue(lines[1].startswith("Pago 2 (00123 · Cliente): "))
        several[1].result = None
        self.assertTrue(payments_message(several).endswith("sin calcular."))

    def test_payment_label(self):
        self.assertEqual(Payment(1).label, "Todos los clientes")
        self.assertEqual(Payment(1, "0419", "Ana").label, "0419 · Ana")
        self.assertEqual(Payment(1, "0419").label, "0419")


def sequential(checks, payments):
    """Un pago tras otro, sin repartir: lo que se hacía antes del reparto conjunto."""
    pool, outcomes = list(checks), []
    for payment in payments:
        outcome = find_payments(pool, [payment], progressive=False)[0]
        outcomes.append(outcome)
        taken = {c.row for c in outcome.result.checks}
        pool = [c for c in pool if c.row not in taken]
    return outcomes


def score(outcomes):
    """(total cubierto, pagos exactos): lo que el reparto maximiza, en ese orden."""
    return (sum(o.result.total for o in outcomes), sum(o.result.transfer == 0 for o in outcomes))


def best_possible(checks, payments):
    """Búsqueda exhaustiva: todas las formas de asignar cada e-cheque a un pago o a ninguno."""
    k, best = len(payments), (-1, -1)
    for assign in itertools.product(range(k + 1), repeat=len(checks)):
        sums = [0] * k
        for check, a in zip(checks, assign):
            if a < k:
                if payments[a].client_id not in (None, check.client_id):
                    break
                sums[a] += check.amount
        else:
            if all(s <= p.target for s, p in zip(sums, payments)):
                best = max(best, (sum(sums), sum(s == p.target for s, p in zip(sums, payments))))
    return best


class SplitTests(unittest.TestCase):
    def assert_valid(self, outcomes, payments):
        rows = [c.row for o in outcomes for c in o.result.checks]
        self.assertEqual(len(rows), len(set(rows)))
        for outcome, payment in zip(outcomes, payments):
            self.assertLessEqual(outcome.result.total, payment.target)
            self.assertTrue(all(payment.client_id in (None, c.client_id) for c in outcome.result.checks))

    def test_splitting_completes_more_payments_than_one_after_another(self):
        checks = checks_for([600, 900, 900, 200, 400])
        payments = [Payment(1200), Payment(1100), Payment(1000)]
        before, after = sequential(checks, payments), find_payments(checks, payments, progressive=False)
        self.assertEqual(sum(o.result.transfer == 0 for o in before), 1)
        self.assertEqual(sum(o.result.transfer == 0 for o in after), 2)
        self.assertGreaterEqual(score(after), score(before))
        self.assert_valid(after, payments)
        self.assertTrue(all(o.result.optimal and o.result.completed for o in after))

    def test_matches_the_exhaustive_optimum_on_random_cases(self):
        rng = random.Random(41)
        improved = 0
        for case in range(120):
            amounts = [rng.choice((100, 200, 300, 400, 500, 600, 700, 800, 900, rng.randrange(1, 90) * 10))
                       for _ in range(rng.randrange(4, 8))]
            checks = [replace(c, client_id=rng.choice("AAB")) for c in checks_for(amounts)]
            payments = [Payment(rng.randrange(2, 16) * 100, rng.choice((None, None, "A", "B")))
                        for _ in range(rng.choice((2, 3)))]
            outcomes = find_payments(checks, payments, progressive=False)
            with self.subTest(case=case, amounts=amounts, targets=[(p.target, p.client_id) for p in payments]):
                self.assertEqual(score(outcomes), best_possible(checks, payments))
                self.assertGreaterEqual(score(outcomes), score(sequential(checks, payments)))
                self.assert_valid(outcomes, payments)
                self.assertTrue(all(o.result.optimal and o.result.completed for o in outcomes))
            improved += score(outcomes) > score(sequential(checks, payments))
        self.assertGreater(improved, 0)

    def test_nothing_to_split_when_every_payment_is_already_exact(self):
        checks, payments = checks_for([500, 400, 300]), [Payment(500), Payment(400)]
        with patch("core._split_search", side_effect=AssertionError("no debía repartir")):
            outcomes = find_payments(checks, payments, progressive=False)
        self.assertEqual([o.result.transfer for o in outcomes], [0, 0])
        self.assertTrue(all(o.result.optimal for o in outcomes))

    def test_progressive_mode_and_single_payment_never_split(self):
        checks = checks_for([600, 900, 900, 200, 400])
        with patch("core._split_search", side_effect=AssertionError("no debía repartir")):
            find_payments(checks, [Payment(1200), Payment(1100), Payment(1000)], progressive=True)
            find_payments(checks, [Payment(1150)], progressive=False)

    def test_exhausted_time_is_reported_not_silent(self):
        checks = checks_for([600, 900, 900, 200, 400])
        payments = [Payment(1200), Payment(1100), Payment(1000)]
        outcomes = find_payments(checks, payments, progressive=False, split_seconds=-1)
        self.assert_valid(outcomes, payments)
        self.assertTrue(all(o.result.completed and not o.result.optimal for o in outcomes))
        self.assertGreaterEqual(score(outcomes), score(sequential(checks, payments)))

    def test_search_status_is_translated_into_each_result(self):
        checks = checks_for([600, 900, 900, 200, 400])
        payments = [Payment(1200), Payment(1100), Payment(1000)]
        for status, completed in (("timeout", True), ("cancelled", False), ("proven", True)):
            with self.subTest(status=status), patch(
                    "core._split_search", side_effect=lambda items, targets, ceilings, seed, *a, s=status: (list(seed), s)):
                outcomes = find_payments(checks, payments, progressive=False)
            self.assertTrue(all(o.result.completed == completed for o in outcomes))
            self.assertTrue(all(o.result.optimal == (status == "proven") for o in outcomes))

    def test_cancel_inside_the_search_keeps_a_valid_split(self):
        rng = random.Random(5)
        amounts = sorted((rng.randrange(1000, 9000) for _ in range(26)), reverse=True)
        total = sum(amounts)
        targets = [int(total * 0.6) + 1, int(total * 0.6) + 3]
        items = [(a, (0, 1)) for a in amounts]
        event = threading.Event()
        event.set()
        best, status = _split_search(items, targets, targets, [-1] * len(items), event, None, time.monotonic() + 60)
        self.assertEqual(status, "cancelled")
        for j, target in enumerate(targets):
            self.assertLessEqual(sum(a for (a, _), who in zip(items, best) if who == j), target)

    def test_progress_reports_the_remaining_total(self):
        rng = random.Random(5)
        amounts = sorted((rng.randrange(1000, 9000) for _ in range(26)), reverse=True)
        total = sum(amounts)
        targets = [int(total * 0.6) + 1, int(total * 0.6) + 3]
        seen = []
        with patch("core.time.monotonic", side_effect=itertools.count(0, 1)):
            _split_search([(a, (0, 1)) for a in amounts], targets, targets, [-1] * len(amounts), None,
                          seen.append, 10 ** 9)
        self.assertTrue(seen and all(m.startswith("Repartiendo entre los pagos") for m in seen))


if __name__ == "__main__":
    unittest.main()
