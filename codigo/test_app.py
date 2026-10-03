import csv
import json
import tempfile
import time
import tkinter as tk
import unittest
from pathlib import Path
from dataclasses import replace
from unittest.mock import patch

from app import Application
from core import ImportResult, Payment, PaymentOutcome, SearchResult, find_combination
from registro import Registry
from test_core import checks_for


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


class AppCase(unittest.TestCase):
    amounts = (30000000, 15000000, 60000000)

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.registry_path = Path(self.temp.name) / "registro.json"
        self.root = tk.Tk()
        self.root.withdraw()
        self.callback_errors = []
        self.root.report_callback_exception = lambda kind, error, trace: self.callback_errors.append(str(error))
        for dialog in ("showwarning", "showerror", "showinfo", "askyesno"):
            guard = patch("app.messagebox." + dialog, side_effect=AssertionError("diálogo inesperado: " + dialog))
            guard.start()
            self.addCleanup(guard.stop)
        self.app = Application(self.root, self.registry_path)
        self.checks = checks_for(list(self.amounts))
        self.app._loaded(("prueba.xlsx", ["Pagos"], ImportResult(self.checks, [], [], "Pagos")))

    def tearDown(self):
        self.root.update_idletasks()
        self.app.close()
        self.temp.cleanup()
        self.assertEqual(self.callback_errors, [])

    def wait_for_search(self):
        deadline = time.monotonic() + 10
        while self.app.busy and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertFalse(self.app.busy)

    def result(self, index=0):
        return self.app.results[index].result

    def check_rows(self, index=0):
        """Filas de e-cheques del pago `index` (cada pago es una fila padre de la tabla)."""
        return self.app.table.get_children(self.app.table.get_children()[index])

    def load(self, records):
        self.app._loaded(("prueba.xlsx", ["Pagos"], ImportResult(records, [], [], "Pagos")))

    def add(self, amount, client=None):
        self.app.target.set(amount)
        if client:
            self.app.client.set(client)
        self.app.add_payment()

    def register(self):
        with patch("app.messagebox.askyesno", return_value=True):
            self.app.register()


class ApplicationTests(AppCase):
    def test_search_transfer_and_invalidate_old_result(self):
        self.app.target.set("500.000")
        self.app.search()
        self.wait_for_search()
        self.assertEqual(self.result().total, 45000000)
        self.assertEqual(self.app.transfer_total.get(), "$ 50.000,00")
        self.assertEqual(len(self.app.table.get_children()), 1)
        self.assertEqual(len(self.check_rows()), 2)
        self.assertEqual(str(self.app.export_button.cget("state")), "normal")
        self.add("250.000")
        self.assertIsNone(self.app.results)
        self.assertFalse(self.app.table.get_children())
        self.assertEqual(str(self.app.copy_button.cget("state")), "disabled")
        self.assertEqual(str(self.app.register_button.cget("state")), "disabled")

    def test_typing_the_next_amount_does_not_wipe_the_current_results(self):
        self.app.target.set("500.000")
        self.app.search()
        self.wait_for_search()
        self.app.target.set("250.000")
        self.app.client.set("Todos los clientes")
        self.assertIsNotNone(self.app.results)
        self.assertEqual(len(self.check_rows()), 2)

    def test_client_filter_and_invalid_target(self):
        other = replace(self.checks[1], client_id="00456", client_name="Otro cliente")
        self.load([self.checks[0], other, self.checks[2]])
        self.add("500.000", next(key for key, value in self.app.clients.items() if value == "00456"))
        self.assertEqual(self.app.payments[0].client_id, "00456")
        self.app.search()
        self.wait_for_search()
        self.assertEqual(self.result().total, 15000000)
        self.app.target.set("-50")
        with patch("app.messagebox.showwarning") as warning:
            self.app.search()
            warning.assert_called_once()
        self.assertEqual(len(self.app.payments), 1)
        self.assertFalse(self.app.busy)

    def test_export_contains_totals_and_preserves_id(self):
        result = find_combination(self.checks, 50000000)
        self.app._found([PaymentOutcome(Payment(50000000), result)])
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "seleccion.csv"
            with patch("app.filedialog.asksaveasfilename", return_value=str(path)):
                self.app.export()
            with path.open(encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.reader(handle, delimiter=";"))
            self.assertIn(["Total e-cheques", "450000,00"], rows)
            self.assertIn(["A completar por transferencia", "50000,00"], rows)
            self.assertIn(["Registrado como utilizado", "No"], rows)
            self.assertIn(["1", "Todos los clientes", "500000,00", "450000,00", "50000,00"], rows)
            self.assertEqual(rows[-1][0], "1")
            self.assertEqual(rows[-1][5], "'00123")

    def test_failed_load_does_not_reuse_previous_file(self):
        self.app.events.put(("error", ("loaded", "Archivo ilegible")))
        with patch("app.messagebox.showerror") as error:
            self.app._poll()
            error.assert_called_once()
        self.assertIsNone(self.app.imported)
        self.assertIsNone(self.app.path)
        self.assertEqual(str(self.app.search_button.cget("state")), "disabled")
        self.assertEqual(str(self.app.add_button.cget("state")), "disabled")

    def test_date_range_filters_and_invalidates_results(self):
        records=[replace(c,date_g=d) for c,d in zip(self.checks,['01/09/2026','08/09/2026','50/09/2026'])]
        self.load(records)
        self.app.target.set('500.000')
        self.app.date_from.set('08/09/2026')
        self.app.date_to.set('08/09/2026')
        self.app.search()
        self.wait_for_search()
        self.assertEqual(self.result().total,15000000)
        self.assertEqual(self.app.search_context['excluded_dates'],1)
        self.assertEqual(self.result().date_field,'G')
        self.app.date_field.set('F')
        self.assertIsNone(self.app.results)
        self.app.date_from.set('20/09/2026')
        with patch('app.messagebox.showwarning') as warning:
            self.app.search()
            warning.assert_called_once()
        self.assertFalse(self.app.busy)

    def test_progressive_and_exact_modes_have_distinct_behavior(self):
        records=checks_for([80000,70000,60000,50000])
        self.load(records)
        self.app.target.set('1100')
        self.app.search()
        self.wait_for_search()
        self.assertEqual(self.result().total,80000)
        self.assertIn('progresiva',self.app.result_title.get())
        self.assertNotIn('detenida',self.app.result_title.get())
        self.app.strategy.set('Mejor suma posible')
        self.assertIsNone(self.app.results)
        self.app.search()
        self.wait_for_search()
        self.assertEqual(self.result().total,110000)
        self.assertTrue(self.result().optimal)

    def test_progressive_export_keeps_selection_order_batches_and_balance(self):
        records=[replace(c,date_g=d) for c,d in zip(checks_for([4000000,90000000,6000000]),
            ['07/09/2026','21/09/2026','14/09/2026'])]
        self.load(records)
        self.app.target.set('1.000.000')
        self.app.search()
        self.wait_for_search()
        self.assertEqual([c.amount for c in self.result().checks],[90000000,6000000,4000000])
        values=[self.app.table.item(row,'values') for row in self.check_rows()]
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
        button=next(w for w in descendants(window) if w.winfo_class()=='TButton' and str(w.cget('text'))=='10')
        button.invoke()
        self.assertEqual(self.app.date_from.get(),'10/09/2025')
        self.app._clear_dates()
        self.assertEqual(self.app.date_from.get(),'')
        self.assertEqual(self.app.date_to.get(),'')


class MultiplePaymentsTests(AppCase):
    """Varios pagos a la vez y registro de e-cheques utilizados."""
    amounts = (50000000, 40000000, 30000000, 20000000)

    def three_payments(self):
        for _ in range(3):
            self.add("500.000")
        self.app.search()
        self.wait_for_search()

    def test_payments_are_listed_and_can_be_removed_or_cleared(self):
        self.add("500.000")
        self.add("250.000")
        values = [self.app.payment_table.item(row, "values") for row in self.app.payment_table.get_children()]
        self.assertEqual([(v[0], v[1], v[2]) for v in values],
                         [("1", "Todos los clientes", "$ 500.000,00"), ("2", "Todos los clientes", "$ 250.000,00")])
        self.assertEqual(self.app.target.get(), "")
        self.assertEqual(self.app.payments_title.get(), "Pagos cargados (2)")
        self.app.payment_table.selection_set(self.app.payment_table.get_children()[0])
        self.app.remove_payment()
        self.assertEqual([p.target for p in self.app.payments], [25000000])
        self.app.clear_payments()
        self.assertEqual(self.app.payments, [])
        self.assertFalse(self.app.payment_table.get_children())

    def test_invalid_amount_is_not_added_and_empty_search_asks_for_a_payment(self):
        self.app.target.set("abc")
        with patch("app.messagebox.showwarning") as warning:
            self.app.add_payment()
            warning.assert_called_once()
        self.assertEqual(self.app.payments, [])
        self.app.target.set("")
        with patch("app.messagebox.showwarning") as warning:
            self.app.search()
            warning.assert_called_once()
        self.assertFalse(self.app.busy)
        self.assertIsNone(self.app.results)

    def test_no_check_is_repeated_between_payments(self):
        self.three_payments()
        chosen = [[c.amount for c in outcome.result.checks] for outcome in self.app.results]
        self.assertEqual(chosen, [[50000000], [40000000], [30000000, 20000000]])
        rows = [c.row for outcome in self.app.results for c in outcome.result.checks]
        self.assertEqual(len(rows), len(set(rows)))
        self.assertEqual(len(self.app.table.get_children()), 3)
        self.assertEqual([len(self.check_rows(i)) for i in range(3)], [1, 1, 2])
        self.assertEqual(self.app.target_total.get(), "$ 1.500.000,00")
        self.assertEqual(self.app.check_total.get(), "$ 1.400.000,00")
        self.assertEqual(self.app.transfer_total.get(), "$ 100.000,00")
        self.assertEqual(self.app.result_title.get(), "3 pagos calculados")
        self.assertIn("Pago 3", self.app.message.get())

    def test_best_sum_splits_checks_across_payments_to_complete_more_of_them(self):
        self.load(checks_for([600, 900, 900, 200, 400]))
        self.app.strategy.set("Mejor suma posible")
        for amount in ("12", "11", "10"):
            self.add(amount)
        self.app.search()
        self.wait_for_search()
        self.assertEqual(sum(o.result.transfer == 0 for o in self.app.results), 2)
        self.assertIn("mejor reparto posible", self.app.result_detail.get())
        rows = [c.row for o in self.app.results for c in o.result.checks]
        self.assertEqual(len(rows), len(set(rows)))
        self.assertEqual(str(self.app.register_button.cget("state")), "normal")

    def test_csv_does_not_call_an_unproven_split_optimal(self):
        result = replace(find_combination(self.checks, 55000000), optimal=False)
        self.assertGreater(result.transfer, 0)
        self.app._found([PaymentOutcome(Payment(55000000), result), PaymentOutcome(Payment(55000000), result)])
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "seleccion.csv"
            with patch("app.filedialog.asksaveasfilename", return_value=str(path)):
                self.app.export()
            with path.open(encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.reader(handle, delimiter=";"))
        self.assertIn(["Estado", "Reparto no demostrado como el mejor (tiempo agotado)"], rows)
        self.assertNotIn(["Estado", "Óptimo confirmado"], rows)
        self.assertIn("No se pudo demostrar", self.app.result_detail.get())

    def test_pending_amount_in_the_field_is_added_when_searching(self):
        self.add("500.000")
        self.app.target.set("400.000")
        self.app.search()
        self.wait_for_search()
        self.assertEqual([p.target for p in self.app.payments], [50000000, 40000000])
        self.assertEqual([c.amount for c in self.result(1).checks], [40000000])

    def test_saving_the_operation_writes_json_and_excludes_those_checks_from_later_searches(self):
        self.three_payments()
        self.assertEqual(str(self.app.register_button.cget("state")), "normal")
        self.assertEqual(str(self.app.register_button.cget("text")), "Guardar operación")
        self.assertFalse(self.registry_path.exists())
        self.app.operation_name.set("  Cobro   de septiembre ")
        with patch("app.messagebox.askyesno") as ask:
            self.app.register()
            ask.assert_not_called()  # guardar es un solo clic; se puede anular desde el registro
        data = json.loads(self.registry_path.read_text(encoding="utf-8"))
        self.assertEqual([p["id"] for p in data["pagos"]], [1, 2, 3])
        self.assertEqual({(p["operacion"], p["nombre_operacion"]) for p in data["pagos"]},
                         {(1, "Cobro de septiembre")})
        self.assertEqual([[c["importe"] for c in p["echeques"]] for p in data["pagos"]],
                         [["500000.00"], ["400000.00"], ["300000.00", "200000.00"]])
        self.assertEqual(data["pagos"][0]["archivo"], "prueba.xlsx")
        self.assertTrue(self.app.registered)
        self.assertEqual(self.app.saved_operation, 1)
        self.assertIn("«Cobro de septiembre» guardada", self.app.save_state.get())
        self.assertEqual(self.app.payments, [])
        self.assertEqual(self.app.available, [])
        self.assertEqual(str(self.app.register_button.cget("state")), "disabled")
        self.assertEqual(str(self.app.saved_button.cget("state")), "normal")
        # Lo calculado sigue a la vista para copiar o exportar.
        self.assertEqual(str(self.app.copy_button.cget("state")), "normal")
        self.assertEqual(len(self.app.table.get_children()), 3)
        self.app.register()
        self.assertEqual(len(Registry.load(self.registry_path).payments), 3)
        self.add("100.000")
        with patch("app.messagebox.showwarning") as warning:
            self.app.search()
            warning.assert_called_once()
            self.assertIn("ya figuran como utilizados", warning.call_args[0][1])
        self.assertIsNone(self.app.results)
        self.assertIsNone(self.app.saved_info)

    def test_the_save_bar_is_always_visible_and_guides_each_step(self):
        self.assertEqual(str(self.app.register_button.cget("state")), "disabled")
        self.assertEqual(str(self.app.register_button.cget("text")), "Guardar operación")
        self.assertTrue(self.app.register_button.winfo_manager())  # empaquetado: aparece en pantalla
        self.assertIn("Calculá los pagos", self.app.save_state.get())
        self.three_payments()
        self.assertIn("presioná «Guardar operación»", self.app.save_state.get())
        self.assertEqual(str(self.app.operation_entry.cget("state")), "normal")
        self.assertEqual(str(self.app.saved_button.cget("state")), "disabled")
        self.add("100.000")  # cambiar los pagos invalida el cálculo: hay que volver a buscar
        self.assertEqual(str(self.app.register_button.cget("state")), "disabled")
        self.assertEqual(str(self.app.operation_entry.cget("state")), "disabled")
        self.assertIn("Calculá los pagos", self.app.save_state.get())
        self.assertFalse(self.registry_path.exists())

    def test_an_unnamed_operation_gets_a_numbered_name_and_each_save_is_a_new_operation(self):
        self.add("500.000")
        self.app.search()
        self.wait_for_search()
        self.app.register()
        self.assertEqual(self.app.saved_info[0], "Operación 1")
        self.add("400.000")
        self.app.search()
        self.wait_for_search()
        self.app.operation_name.set("Segundo cobro")
        self.app.register()
        operations = Registry.load(self.registry_path).operations()
        self.assertEqual([(o.id, o.label) for o in operations], [(1, "Operación 1"), (2, "Segundo cobro")])
        self.assertIn("2 operaciones", self.app.registry_info.get())

    def test_a_search_alone_never_burns_checks(self):
        self.three_payments()
        self.assertFalse(self.registry_path.exists())
        self.app.search()
        self.wait_for_search()
        self.assertEqual([c.amount for c in self.result(0).checks], [50000000])

    def test_registered_checks_stay_blocked_in_a_new_listing_with_other_rows(self):
        self.add("500.000")
        self.app.search()
        self.wait_for_search()
        self.register()
        shuffled = [replace(c, row=c.row + 100) for c in reversed(self.checks)]
        self.load(shuffled)
        self.assertEqual(sorted(c.amount for c in self.app.available), [20000000, 30000000, 40000000])
        self.assertIn("ya utilizados", self.app.file_detail.get())
        self.add("500.000")
        self.app.search()
        self.wait_for_search()
        self.assertEqual([c.amount for c in self.result().checks], [40000000])

    def test_second_batch_continues_where_the_first_stopped(self):
        self.add("500.000")
        self.add("400.000")
        self.app.search()
        self.wait_for_search()
        self.register()
        self.add("500.000")
        self.app.search()
        self.wait_for_search()
        self.assertEqual([c.amount for c in self.result().checks], [30000000, 20000000])
        self.register()
        self.assertEqual([len(p.checks) for p in Registry.load(self.registry_path).payments], [1, 1, 2])

    def test_checks_registered_elsewhere_in_the_meantime_are_a_conflict(self):
        self.add("500.000")
        self.app.search()
        self.wait_for_search()
        rival = Registry.load(self.registry_path)
        rival.add([PaymentOutcome(Payment(500), SearchResult([self.checks[0]], 500, True, 0.0))], "otro.xlsx", "Pagos")
        rival.save()
        with patch("app.messagebox.showerror") as error:
            self.register()
            error.assert_called_once()
        self.assertEqual(len(Registry.load(self.registry_path).payments), 1)
        self.assertFalse(self.app.registered)

    def test_unreadable_registry_blocks_searching_and_is_never_overwritten(self):
        self.registry_path.write_text("{esto no es json", encoding="utf-8")
        self.load(self.checks)
        self.assertTrue(self.app.registry_info.get().startswith("⚠"))
        self.assertEqual(self.app.available, [])
        self.add("500.000")
        with patch("app.messagebox.showwarning") as warning:
            self.app.search()
            warning.assert_called_once()
        self.assertFalse(self.app.busy)
        self.assertEqual(self.registry_path.read_text(encoding="utf-8"), "{esto no es json")

    def test_stopped_batch_cannot_be_registered_and_marks_uncalculated_payments(self):
        first = find_combination(self.checks, 50000000)
        outcomes = [PaymentOutcome(Payment(50000000), first), PaymentOutcome(Payment(40000000), None)]
        self.app._found(outcomes)
        self.assertEqual(str(self.app.register_button.cget("state")), "disabled")
        self.assertEqual(str(self.app.copy_button.cget("state")), "normal")
        self.assertIn("detenida", self.app.result_title.get())
        self.assertIn("sin calcular", self.app.message.get())
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "seleccion.csv"
            with patch("app.filedialog.asksaveasfilename", return_value=str(path)):
                self.app.export()
            with path.open(encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.reader(handle, delimiter=";"))
        self.assertIn(["2", "Todos los clientes", "400000,00", "sin calcular", ""], rows)

    def open_registry(self, **options):
        self.app.show_registry(**options)
        window = [w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel)][-1]
        operations_view, checks_view = [w for w in descendants(window) if w.winfo_class() == "Treeview"]
        void = next(w for w in descendants(window) if w.winfo_class() == "TButton" and "Anular" in str(w.cget("text")))
        return window, operations_view, checks_view, void

    def test_registry_window_identifies_each_operation_and_clicking_shows_its_checks(self):
        self.add("500.000")
        self.add("400.000")
        self.app.search()
        self.wait_for_search()
        self.app.operation_name.set("Cobro A")
        self.app.register()
        self.add("200.000")
        self.app.search()
        self.wait_for_search()
        self.app.register()
        window, operations_view, checks_view, _ = self.open_registry()
        self.assertEqual(operations_view.get_children(), ("op2", "op1"))  # la más reciente primero
        first = operations_view.item("op1")
        self.assertEqual(first["text"], "#1 · Cobro A")
        self.assertEqual(first["values"][1], "2 pagos")
        self.assertEqual(first["values"][5], "Registrada")
        self.assertEqual(operations_view.item("op2")["text"], "#2 · Operación 2")
        self.assertEqual(operations_view.get_children("op1"), ("pay1", "pay2"))
        operations_view.selection_set("op1")
        self.root.update()
        self.assertEqual(len(checks_view.get_children()), 2)  # todos los e-cheques de la operación
        operations_view.selection_set("pay2")
        self.root.update()
        self.assertEqual([checks_view.item(i, "values")[5] for i in checks_view.get_children()], ["$ 400.000,00"])
        window.destroy()

    def test_saved_button_opens_the_registry_on_that_operation(self):
        self.add("500.000")
        self.app.search()
        self.wait_for_search()
        self.app.operation_name.set("Cobro A")
        self.app.register()
        self.add("400.000")
        self.app.search()
        self.wait_for_search()
        self.app.register()
        self.app.saved_operation = 1  # volver a la primera operación guardada
        self.app.show_saved()
        window = [w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel)][-1]
        operations_view = next(w for w in descendants(window) if w.winfo_class() == "Treeview")
        self.assertEqual(operations_view.selection(), ("op1",))
        window.destroy()

    def test_voiding_an_operation_frees_all_its_checks_and_keeps_the_history(self):
        self.three_payments()
        self.app.register()
        window, operations_view, checks_view, void = self.open_registry()
        operations_view.selection_set("op1")
        self.root.update()
        with patch("app.messagebox.askyesno", return_value=True):
            void.invoke()
        self.assertEqual(len(self.app.available), 4)
        voided = Registry.load(self.registry_path)
        self.assertTrue(all(p.voided for p in voided.payments))
        self.assertEqual(len(voided.payments), 3)
        self.assertEqual(operations_view.item("op1")["values"][5], "Anulada")
        window.destroy()
        self.add("500.000")
        self.app.search()
        self.wait_for_search()
        self.assertEqual([c.amount for c in self.result().checks], [50000000])

    def test_voiding_one_payment_frees_only_its_checks(self):
        self.three_payments()
        self.app.register()
        window, operations_view, checks_view, void = self.open_registry()
        operations_view.selection_set("pay3")
        self.root.update()
        with patch("app.messagebox.askyesno", return_value=True):
            void.invoke()
        self.assertEqual(sorted(c.amount for c in self.app.available), [20000000, 30000000])
        self.assertEqual(operations_view.item("op1")["values"][5], "Parcialmente anulada")
        self.assertTrue(operations_view.item("pay3")["values"][5].startswith("Anulado"))
        window.destroy()
        self.add("500.000")
        self.app.search()
        self.wait_for_search()
        self.assertEqual([c.amount for c in self.result().checks], [30000000, 20000000])

    def test_registry_window_requires_a_selection_to_void(self):
        self.app.show_registry()
        window = next(w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel))
        button = next(w for w in descendants(window) if w.winfo_class() == "TButton" and "Anular" in str(w.cget("text")))
        with patch("app.messagebox.showinfo") as info, patch("app.messagebox.askyesno") as ask:
            button.invoke()
            info.assert_called_once()
            ask.assert_not_called()
        window.destroy()


if __name__ == "__main__":
    unittest.main()
