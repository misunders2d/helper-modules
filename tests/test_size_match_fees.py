"""Focused fee checks; load pure calculations without GUI/cloud initialization."""
import ast
from pathlib import Path
import subprocess
import unittest

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]


def calculations(source):
    tree = ast.parse(source)
    nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef)
             and n.name in {'get_shipping_weight', 'get_size_tier', 'get_fulfillment_fee'}]
    constants = {n.targets[0].id: ast.literal_eval(n.value) for n in tree.body
                 if isinstance(n, ast.Assign) and isinstance(n.targets[0], ast.Name)
                 and n.targets[0].id.isupper() and isinstance(n.value, ast.Constant)}
    ns = dict(constants, np=np, pd=pd, today=pd.Timestamp('2026-11-06'))
    exec(compile(ast.Module(body=nodes, type_ignores=[]), '<calculations>', 'exec'), ns)
    return ns


class FeeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.calc = calculations((ROOT / 'utils/size_match.py').read_text())

    def row(self, weight=3.3, price=40, tier='Large standard size'):
        return pd.DataFrame({'size_tier': [tier], 'shipping_weight': [weight],
                             'individual weight lbs': [weight], 'price': [price]})

    def fee(self, **kw):
        return self.calc['get_fulfillment_fee'](self.row(**kw), schedule='2027_non_peak').fba_fee.iloc[0]

    def test_legacy_regression(self):
        old = calculations(subprocess.check_output(['git', 'show', '2af9e09:utils/size_match.py'], cwd=ROOT, text=True))
        rows = pd.concat([self.row(w, p, t) for t, weights in [
            ('Small standard size', [.125, .25, .75, 1]),
            ('Large standard size', [.25, 1, 1.5, 3, 3.3, 20]),
            ('Small bulky', [25]), ('Large bulky', [40])]
            for w in weights for p in [9.99, 10, 50, 50.01]], ignore_index=True)
        for date in ['2025-10-14', '2025-10-15', '2026-01-14', '2026-01-15', '2026-10-14', '2026-10-15']:
            old['today'] = pd.Timestamp(date)
            expected = old['get_fulfillment_fee'](rows.copy())
            actual = self.calc['get_fulfillment_fee'](rows.copy(), as_of=date, schedule="legacy")
            pd.testing.assert_frame_equal(expected, actual)

    def test_pdf_examples_and_price_boundaries(self):
        # PDF p12 mobile case and p13 T-shirt (non-apparel row illustration).
        self.assertEqual(self.fee(weight=2.88/16, price=9, tier='Small standard size'), 2.72)
        self.assertEqual(self.fee(weight=10/16, price=40), 4.26)
        for p, expected in [(9.99, 6.31), (10, 7.32), (50, 7.32), (50.01, 7.89)]:
            self.assertEqual(self.fee(weight=3.3, price=p), expected)

    def test_interval_boundaries(self):
        for w, expected in [(3, 6.85), (3.00001, 7.24), (3.25, 7.24), (3.25001, 7.32), (20, 12.60)]:
            self.assertEqual(self.fee(weight=w), expected)
        self.assertEqual(self.fee(weight=2/16, tier='Small standard size'), 3.75)
        self.assertEqual(self.fee(weight=2.00001/16, tier='Small standard size'), 3.78)

    def test_effective_dates_and_unpublished_peak(self):
        f = self.calc['get_fulfillment_fee']
        self.assertAlmostEqual(f(self.row(), as_of='2027-01-14').fba_fee.iloc[0], 7.67)
        self.assertEqual(f(self.row(), as_of='2027-01-15').fba_fee.iloc[0], 7.32)
        self.assertEqual(f(self.row(), as_of='2027-10-14').fba_fee.iloc[0], 7.32)
        self.assertTrue(pd.isna(f(self.row(), schedule='2027_non_peak').fba_peak_fee.iloc[0]))
        for kw in [dict(as_of='2027-10-15'), dict(schedule='2027_peak'), dict(schedule='wrong')]:
            with self.assertRaises(ValueError):
                f(self.row(), **kw)
        self.assertEqual(f(self.row(), schedule='2026_peak').fba_fee.iloc[0], 7.67)
        for date, expected in [('2026-10-14', 7.13), ('2026-10-15', 7.67), ('2026-01-15', 7.13)]:
            self.assertEqual(f(self.row(), as_of=date).fba_fee.iloc[0], expected)

    def test_active_mixed_frame_preserves_unrelated_tiers(self):
        f = self.calc['get_fulfillment_fee']
        frame = pd.concat([self.row(), self.row(weight=25, tier='Small bulky')], ignore_index=True)
        result = f(frame.copy(), as_of='2026-11-06')
        legacy = f(frame.copy(), as_of='2026-11-06', schedule='legacy')
        self.assertEqual(result.fba_fee.iloc[0], 7.67)
        self.assertEqual(result.fba_fee.iloc[1], legacy.fba_fee.iloc[1])
        self.assertEqual(result.fulfillment_schedule.iloc[1], 'legacy (non-standard size)')
        historical = f(frame.copy(), as_of='2025-11-06')
        pd.testing.assert_frame_equal(historical, f(frame.copy(), as_of='2025-11-06', schedule='legacy'))

    def test_package_weight_and_invalid_rows(self):
        f = self.calc['get_fulfillment_fee']
        df = pd.DataFrame({'l':[11.5], 'w':[9.5], 'h':[2], 'individual weight lbs':[2.97], 'price':[36.97]})
        df = self.calc['get_shipping_weight'](df)
        self.assertAlmostEqual(df.dim_weight.iloc[0], 218.5/139)
        df = self.calc['get_size_tier'](df)
        self.assertEqual(df.size_tier.iloc[0], 'Large standard size')
        for row in [self.row(weight=0), self.row(weight=20.01), self.row(price=np.nan), self.row(tier='Small bulky'), self.row(weight=1.01, tier='Small standard size')]:
            with self.assertRaises(ValueError):
                f(row, schedule='2027_non_peak')


if __name__ == '__main__':
    unittest.main()
