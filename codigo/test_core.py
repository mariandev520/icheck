import random
import tempfile
import threading
import unittest
import zipfile
from dataclasses import replace
from datetime import date
from pathlib import Path
from xml.sax.saxutils import escape

from core import (Check, DataError, batch_label, check_date, customer_message, filter_checks,
                  find_combination, find_progressive, list_sheets, money, numeric_cents,
                  parse_amount, parse_date, read_excel, week_start)


def checks_for(values):
    return [Check(i + 2, str(i), "01/09/2025", "15/09/2025", "00123", "Cliente", value, "Recibo")
            for i, value in enumerate(values)]


def workbook(path, rows, date1904=False):
    cells = []
    for number, data in enumerate(rows, 1):
        items = []
        for column, value in data.items():
            if isinstance(value, tuple):
                text, kind = value
            elif isinstance(value, (int, float)):
                text, kind = str(value), "n"
            else:
                text, kind = str(value), "inlineStr"
            if kind == "inlineStr":
                payload = "<is><t>{}</t></is>".format(escape(text))
            elif kind == "formula":
                kind = "n"
                payload = "<f>1+1</f><v>{}</v>".format(escape(text))
            else:
                payload = "<v>{}</v>".format(escape(text))
            items.append('<c r="{}{}" t="{}">{}</c>'.format(column, number, kind, payload))
        cells.append('<row r="{}">{}</row>'.format(number, "".join(items)))
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", '<workbook xmlns="{}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="{}"/><sheets><sheet name="Pagos" sheetId="1" r:id="rId1"/></sheets></workbook>'.format(ns, int(date1904)))
        archive.writestr("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        archive.writestr("xl/worksheets/sheet1.xml", '<worksheet xmlns="{}"><sheetData>{}</sheetData></worksheet>'.format(ns, "".join(cells)))


class MoneyTests(unittest.TestCase):
    def test_argentine_and_decimal_inputs(self):
        for value, expected in (("500.000", 50000000), ("$ 500.000,50", 50000050),
                                ("500000.50", 50000050), ("500000,5", 50000050),
                                ("1.234.567,89", 123456789), ("0,01", 1),
                                ("ARS 100", 10000), ("100.25", 10025)):
            self.assertEqual(parse_amount(value), expected)

    def test_reject_ambiguous_or_invalid_input(self):
        for value in ("", "-50", "0", "500,000", "50.00.0", "NaN", "1e6", "50/08/2025", "1234.567"):
            with self.subTest(value=value), self.assertRaises(DataError):
                parse_amount(value)

    def test_excel_binary_precision(self):
        self.assertEqual(numeric_cents("1425520.6399999999"), 142552064)
        self.assertEqual(numeric_cents("596200.05000000005"), 59620005)
        self.assertEqual(numeric_cents("100000"), 10000000)
        for value in ("1.005", "NaN", "Infinity", "-100", "0", "0.001"):
            with self.subTest(value=value), self.assertRaises(DataError):
                numeric_cents(value)

    def test_format(self):
        self.assertEqual(money(50000050), "$ 500.000,50")
        self.assertEqual(money(1), "$ 0,01")
        self.assertEqual(money(0), "$ 0,00")


class SolverTests(unittest.TestCase):
    def assert_solution(self, values, target, expected, memory_mb=160):
        items = checks_for(values)
        result = find_combination(items, target, memory_mb=memory_mb)
        self.assertEqual(result.total, expected)
        self.assertTrue(result.optimal)
        self.assertLessEqual(result.total, target)
        self.assertEqual(result.transfer, target - result.total)
        self.assertEqual(len(set(check.row for check in result.checks)), len(result.checks))
        self.assertTrue(all(check in items for check in result.checks))
        return result

    def test_greedy_is_not_sufficient(self):
        for memory in (0, 160):
            self.assert_solution([800, 700, 600, 500], 1100, 1100, memory)
            self.assert_solution([830, 715, 609, 506], 1116, 1115, memory)

    def test_transfer_and_no_overshoot(self):
        result = self.assert_solution([30000000, 15000000, 60000000], 50000000, 45000000)
        self.assertIn("$ 50.000,00", customer_message(result))

    def test_no_reuse_and_same_amount_distinct_checks(self):
        self.assert_solution([30000], 60000, 30000)
        self.assert_solution([30000, 30000], 60000, 60000)
        self.assert_solution([99, 99, 99, 250], 199, 198)

    def test_no_fit_all_fit_and_empty(self):
        self.assert_solution([1000, 2000], 500, 0)
        self.assert_solution([1000, 2000], 4000, 3000)
        self.assert_solution([], 500, 0)

    def test_cent_accuracy(self):
        self.assert_solution([10001, 19999, 40000], 30000, 30000)
        self.assert_solution([20003, 20004], 40006, 20004)
        self.assert_solution([200, 400, 800], 1001, 1000)

    def test_invalid_values(self):
        with self.assertRaises(DataError):
            find_combination(checks_for([0]), 100)
        with self.assertRaises(DataError):
            find_combination([], 0)

    def test_cancellation_is_labeled(self):
        event = threading.Event()
        event.set()
        result = find_combination(checks_for([830, 715, 609, 506]), 1116, event)
        self.assertFalse(result.optimal)
        self.assertLessEqual(result.total, 1116)

    def test_both_algorithms_against_exhaustive_search(self):
        rng = random.Random(9817)
        for iteration in range(160):
            values = [rng.randrange(1, 800) for _ in range(rng.randrange(1, 13))]
            target = rng.randrange(1, sum(values) + 500)
            reachable = {0}
            for value in values:
                reachable |= {total + value for total in list(reachable) if total + value <= target}
            expected = max(reachable)
            for memory in (0, 160):
                with self.subTest(iteration=iteration, memory=memory):
                    self.assert_solution(values, target, expected, memory)

    def test_checkpoint_reconstruction_many_blocks(self):
        rng = random.Random(782)
        values = [rng.randrange(100, 1000) * 100 for _ in range(190)]
        reachable = {0}
        target = 340055
        for value in values:
            reachable |= {total + value for total in list(reachable) if total + value <= target}
        self.assert_solution(values, target, max(reachable))


class ProgressiveTests(unittest.TestCase):
    def dated(self, amounts, dates):
        return [replace(c, date_g=d) for c, d in zip(checks_for(amounts), dates)]

    def test_exact_single_wins_over_multiple_checks(self):
        result = find_progressive(self.dated([900, 100, 1000, 1000],
            ['21/09/2026', '14/09/2026', '07/09/2026', '21/09/2026']), 1000)
        self.assertEqual([c.row for c in result.checks], [5])
        self.assertTrue(result.completed)
        self.assertTrue(result.optimal)

    def test_closest_then_previous_weeks_and_descending_amounts(self):
        records = self.dated([90000000, 6000000, 4000000],
                             ['21/09/2026', '14/09/2026', '07/09/2026'])
        result = find_progressive(records, 100000000)
        self.assertEqual([c.amount for c in result.checks], [90000000,6000000,4000000])
        self.assertEqual(result.transfer, 0)
        self.assertEqual(result.date_field, 'G')

    def test_exact_remaining_has_priority_over_previous_week(self):
        records = self.dated([900, 70, 100], ['21/09/2026','14/09/2026','28/09/2026'])
        result = find_progressive(records,1000)
        self.assertEqual([c.amount for c in result.checks],[900,100])

    def test_previous_week_preference_is_optional_and_not_a_global_optimum(self):
        records = self.dated([800, 150, 100, 60],
            ['21/09/2026','21/09/2026','14/09/2026','07/09/2026'])
        preferred = find_progressive(records,1000)
        by_amount = find_progressive(records,1000,prefer_previous=False)
        self.assertEqual([c.amount for c in preferred.checks],[800,100,60])
        self.assertEqual([c.amount for c in by_amount.checks],[800,150])
        self.assertTrue(preferred.completed)
        self.assertFalse(preferred.optimal)
        # Mayor a menor es un criterio distinto de maximizar la suma.
        records = checks_for([800,700,600,500])
        self.assertEqual(find_progressive(records,1100).total,800)
        self.assertEqual(find_combination(records,1100).total,1100)

    def test_fallback_for_missing_previous_week_and_invalid_dates(self):
        records = self.dated([800,120,60],['21/09/2026','07/09/2026','fecha inválida'])
        result = find_progressive(records,1000)
        self.assertEqual(result.total,980)
        self.assertEqual(len(result.checks),3)
        self.assertEqual(batch_label(result.checks[-1]),'Sin fecha válida')

    def test_weeks_cross_year_and_weekend_boundary(self):
        self.assertEqual(week_start(date(2027,1,3)),date(2026,12,28))
        self.assertEqual(week_start(date(2027,1,4)),date(2027,1,4))
        records = self.dated([800,100,90],['04/01/2027','03/01/2027','27/12/2026'])
        self.assertEqual([c.amount for c in find_progressive(records,1000).checks],[800,100,90])

    def test_dates_are_inclusive_and_choose_the_selected_column(self):
        records = self.dated([500,400,300,200],['14/09/2026','20/09/2026','21/09/2026','50/09/2026'])
        filtered, invalid = filter_checks(records,'G',date(2026,9,14),date(2026,9,20))
        self.assertEqual([c.amount for c in filtered],[500,400])
        self.assertEqual(invalid,1)
        self.assertEqual(filter_checks(records,'F',date(2025,9,1),date(2025,9,1))[0],records)
        self.assertEqual(filter_checks(records)[0],records)
        self.assertIsNone(check_date(records[-1]))
        with self.assertRaises(DataError):
            filter_checks(records,'G',date(2026,10,1),date(2026,9,1))
        with self.assertRaises(DataError):
            parse_date('31/02/2026')
        self.assertEqual(parse_date('29/02/2024'),date(2024,2,29))

    def test_cancellation_and_no_overshoot(self):
        event=threading.Event()
        event.set()
        result=find_progressive(checks_for([900,600,400]),1000,event)
        self.assertFalse(result.completed)
        self.assertFalse(result.optimal)
        self.assertEqual(result.total,0)
        self.assertEqual(find_progressive(checks_for([1100]),1000).checks,[])

    def test_random_cases_never_reuse_or_exceed_and_are_descending(self):
        rng=random.Random(178)
        for _ in range(100):
            records=self.dated([rng.randrange(1,1000) for _ in range(100)],
                [rng.choice(['01/09/2026','09/09/2026','15/09/2026','28/09/2026','']) for _ in range(100)])
            target=rng.randrange(1,10000)
            result=find_progressive(records,target)
            amounts=[c.amount for c in result.checks]
            self.assertEqual(amounts,sorted(amounts,reverse=True))
            self.assertLessEqual(result.total,target)
            self.assertEqual(len({c.row for c in result.checks}),len(result.checks))
            self.assertTrue(all(c in records for c in result.checks))
            self.assertTrue(result.completed)

    def test_twenty_thousand_checks(self):
        result=find_progressive(checks_for([100]*20000),2000001)
        self.assertEqual(len(result.checks),20000)
        self.assertEqual(result.transfer,1)
        self.assertTrue(result.optimal)


class ImportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name) / "test.xlsx"

    def tearDown(self):
        self.temp.cleanup()

    def test_marked_dates_ids_and_integer_amounts(self):
        workbook(self.path, [{"F": 45891, "G": "50/08/2025", "J": "0419", "R": "Cliente", "V": 502511}])
        result = read_excel(self.path)
        self.assertEqual(list_sheets(self.path), ["Pagos"])
        self.assertEqual(len(result.checks), 1)
        check = result.checks[0]
        self.assertEqual(check.date_f, "22/08/2025")
        self.assertEqual(check.date_g, "50/08/2025")
        self.assertEqual(check.client_id, "0419")
        self.assertEqual(check.amount, 50251100)
        self.assertEqual(len(result.notes), 2)
        self.assertFalse(result.rejected)

    def test_headers_blank_rows_bad_amount_and_missing_id(self):
        workbook(self.path, [{"J": "ID cliente", "V": "Importe"}, {},
                             {"J": "00123", "V": "500.000,50", "F": "01/09/2025", "G": "02/09/2025"},
                             {"J": "00123", "V": "inválido"}, {"V": 100},
                             {"J": "00123", "V": -100}])
        result = read_excel(self.path)
        self.assertEqual(len(result.checks), 1)
        self.assertEqual(len(result.rejected), 3)
        self.assertEqual(result.checks[0].row, 3)
        self.assertEqual(result.checks[0].amount, 50000050)

    def test_repeated_dates_id_and_amount_are_distinct_checks(self):
        workbook(self.path, [{"C": "0001", "F": 45891, "G": 45897, "J": "00123", "V": 300},
                             {"C": "0002", "F": 45891, "G": 45897, "J": "00123", "V": 300}])
        result = find_combination(read_excel(self.path).checks, 60000)
        self.assertEqual(len(result.checks), 2)
        self.assertEqual(result.total, 60000)

    def test_numeric_id_1904_dates_and_formula_cache(self):
        workbook(self.path, [{"F": 1, "G": 2, "J": 123, "V": ("2", "formula")}], date1904=True)
        result = read_excel(self.path)
        self.assertEqual(result.checks[0].client_id, "00123")
        self.assertEqual(result.checks[0].date_f, "02/01/1904")
        self.assertEqual(result.checks[0].amount, 200)
        self.assertIn("fórmula", result.notes[0])

    def test_bad_file_and_missing_sheet(self):
        self.path.write_text("no es un excel")
        with self.assertRaises(DataError):
            read_excel(self.path)
        workbook(self.path, [])
        with self.assertRaises(DataError):
            read_excel(self.path, "Otra")


if __name__ == "__main__":
    unittest.main()
