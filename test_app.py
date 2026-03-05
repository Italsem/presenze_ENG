import os
import tempfile
import unittest

import app


class PresenzeTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        app.DB_NAME = os.path.join(self.tmp.name, "test_presenze.db")
        app.db_init()

    def tearDown(self):
        self.tmp.cleanup()

    def test_calc_work_minutes_with_break(self):
        self.assertEqual(app.calc_work_minutes("08:30", "12:30", "13:30", "17:30"), 480)

    def test_calc_work_minutes_without_break(self):
        self.assertEqual(app.calc_work_minutes("09:00", "", "", "17:00"), 480)

    def test_validate_work_times(self):
        ok, _ = app.validate_work_times("08:00", "12:00", "13:00", "17:00")
        self.assertTrue(ok)
        ok, _ = app.validate_work_times("08:00", "12:00", "", "17:00")
        self.assertFalse(ok)
        ok, _ = app.validate_work_times("18:00", "", "", "17:00")
        self.assertFalse(ok)

    def test_presence_insert_and_month_stats(self):
        app.db_add_employee("Mario", "Rossi", 20)
        emp_id = app.db_list_employees()[0][0]

        app.db_add_presence(emp_id, "2026-01-02", "Lavoro", "08:00", "12:00", "13:00", "17:00", 480, "")
        app.db_add_presence(emp_id, "2026-01-03", "Ferie", "", "", "", "", 0, "")

        tot, giorni, media, ferie = app.db_month_stats(emp_id, "2026-01")
        self.assertEqual(tot, 480)
        self.assertEqual(giorni, 1)
        self.assertEqual(media, 480)
        self.assertEqual(ferie, 1)
        self.assertEqual(app.db_year_ferie(emp_id, 2026), 1)

    def test_export_pdf(self):
        out = os.path.join(self.tmp.name, "report.pdf")
        rows = [(1, "2026-01-02", "Lavoro", "08:00", "12:00", "13:00", "17:00", 480, "ok")]
        app.export_month_pdf(out, "Rossi Mario", "2026-01", rows, (480, 1, 480, 0), (20, 0, 20))
        self.assertTrue(os.path.exists(out))
        self.assertGreater(os.path.getsize(out), 100)


if __name__ == "__main__":
    unittest.main()
