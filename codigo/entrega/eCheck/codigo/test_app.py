import csv
import tempfile
import time
import tkinter as tk
import unittest
from pathlib import Path
from dataclasses import replace
from unittest.mock import patch

from app import Application
from core import ImportResult, find_combination
from test_core import checks_for


class ApplicationTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.callback_errors = []
        self.root.report_callback_exception = lambda kind, error, trace: self.callback_errors.append(str(error))
        self.app = Application(self.root)
        self.checks = checks_for([30000000, 15000000, 60000000])
        self.app._loaded(("prueba.xlsx", ["Pagos"], ImportResult(self.checks, [], [], "Pagos")))

    def tearDown(self):
        self.root.update_idletasks()
        self.app.close()
        self.assertEqual(self.callback_errors, [])

    def wait_for_search(self):
        deadline = time.monotonic() + 10
        while self.app.busy and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertFalse(self.app.busy)

    def test_search_transfer_and_invalidate_old_result(self):
        self.app.target.set("500.000")
        self.app.search()
        self.wait_for_search()
        self.assertEqual(self.app.result.total, 45000000)
        self.assertEqual(self.app.transfer_total.get(), "$ 50.000,00")
        self.assertEqual(len(self.app.table.get_children()), 2)
        self.assertEqual(str(self.app.export_button.cget("state")), "normal")
        self.app.target.set("250.000")
        self.assertIsNone(self.app.result)
        self.assertFalse(self.app.table.get_children())
        self.assertEqual(str(self.app.copy_button.cget("state")), "disabled")

    def test_client_filter_and_invalid_target(self):
        from dataclasses import replace
        other = replace(self.checks[1], client_id="00456", client_name="Otro cliente")
        records = [self.checks[0], other, self.checks[2]]
        self.app._loaded(("prueba.xlsx", ["Pagos"], ImportResult(records, [], [], "Pagos")))
        self.app.client.set(next(key for key, value in self.app.clients.items() if value == "00456"))
        self.app.target.set("500.000")
        self.app.search()
        self.wait_for_search()
        self.assertEqual(self.app.result.total, 15000000)
        self.app.target.set("-50")
        with patch("app.messagebox.showwarning") as warning:
            self.app.search()
            warning.assert_called_once()
        self.assertIsNone(self.app.result)

    def test_export_contains_totals_and_preserves_id(self):
        self.app._found(find_combination(self.checks, 50000000))
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "seleccion.csv"
            with patch("app.filedialog.asksaveasfilename", return_value=str(path)):
                self.app.export()
            with path.open(encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.reader(handle, delimiter=";"))
            self.assertIn(["Total e-cheques", "450000,00"], rows)
            self.assertIn(["A completar por transferencia", "50000,00"], rows)
            self.assertEqual(rows[-1][4], "'00123")

    def test_failed_load_does_not_reuse_previous_file(self):
        self.app.events.put(("error", ("loaded", "Archivo ilegible")))
        with patch("app.messagebox.showerror") as error:
            self.app._poll()
            error.assert_called_once()
        self.assertIsNone(self.app.imported)
        self.assertIsNone(self.app.path)
        self.assertEqual(str(self.app.search_button.cget("state")), "disabled")

    def test_date_range_filters_and_invalidates_results(self):
        records=[replace(c,date_g=d) for c,d in zip(self.checks,['01/09/2026','08/09/2026','50/09/2026'])]
        self.app._loaded(('prueba.xlsx',['Pagos'],ImportResult(records,[],[],'Pagos')))
        self.app.target.set('500.000')
        self.app.date_from.set('08/09/2026')
        self.app.date_to.set('08/09/2026')
        self.app.search()
        self.wait_for_search()
        self.assertEqual(self.app.result.total,15000000)
        self.assertEqual(self.app.search_context['excluded_dates'],1)
        self.assertEqual(self.app.result.date_field,'G')
        self.app.date_field.set('F')
        self.assertIsNone(self.app.result)
        self.app.date_from.set('20/09/2026')
        with patch('app.messagebox.showwarning') as warning:
            self.app.search()
            warning.assert_called_once()
        self.assertFalse(self.app.busy)

    def test_progressive_and_exact_modes_have_distinct_behavior(self):
        records=checks_for([80000,70000,60000,50000])
        self.app._loaded(('prueba.xlsx',['Pagos'],ImportResult(records,[],[],'Pagos')))
        self.app.target.set('1100')
        self.app.search()
        self.wait_for_search()
        self.assertEqual(self.app.result.total,80000)
        self.assertIn('progresiva',self.app.result_title.get())
        self.assertNotIn('detenida',self.app.result_title.get())
        self.app.strategy.set('Mejor suma posible')
        self.assertIsNone(self.app.result)
        self.app.search()
        self.wait_for_search()
        self.assertEqual(self.app.result.total,110000)
        self.assertTrue(self.app.result.optimal)

    def test_progressive_export_keeps_selection_order_batches_and_balance(self):
        records=[replace(c,date_g=d) for c,d in zip(checks_for([4000000,90000000,6000000]),
            ['07/09/2026','21/09/2026','14/09/2026'])]
        self.app._loaded(('prueba.xlsx',['Pagos'],ImportResult(records,[],[],'Pagos')))
        self.app.target.set('1.000.000')
        self.app.search()
        self.wait_for_search()
        self.assertEqual([c.amount for c in self.app.result.checks],[90000000,6000000,4000000])
        values=[self.app.table.item(row,'values') for row in self.app.table.get_children()]
        self.assertEqual([v[-1] for v in values],['$ 100.000,00','$ 40.000,00','$ 0,00'])
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/'seleccion.csv'
            with patch('app.filedialog.asksaveasfilename',return_value=str(path)):
                self.app.export()
            with path.open(encoding='utf-8-sig',newline='') as handle:
                rows=list(csv.reader(handle,delimiter=';'))
            self.assertIn(['Estado','Selección progresiva'],rows)
            self.assertEqual([r[-1] for r in rows[-3:]],['100000,00','40000,00','0,00'])

    def test_calendar_selects_date_and_clear_restores_open_range(self):
        self.app._open_calendar(self.app.date_from)
        window=next(w for w in self.root.winfo_children() if isinstance(w,tk.Toplevel))
        def descendants(widget):
            for child in widget.winfo_children():
                yield child
                yield from descendants(child)
        button=next(w for w in descendants(window) if w.winfo_class()=='TButton' and str(w.cget('text'))=='10')
        button.invoke()
        self.assertEqual(self.app.date_from.get(),'10/09/2025')
        self.app._clear_dates()
        self.assertEqual(self.app.date_from.get(),'')
        self.assertEqual(self.app.date_to.get(),'')


if __name__ == "__main__":
    unittest.main()
