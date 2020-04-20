import unittest

import os
import pandas as pd
import sys

sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(__file__))))
from src.aggregator import aggregate


class TestAggregator(unittest.TestCase):
    def test_aggregate(self):
        for file in os.listdir(os.path.join('data', 'input')):
            with self.subTest(file=file):
                df = aggregate(os.path.join('data', 'input', file))
                df_expected = pd.read_excel(os.path.join('data', 'expected', 'Res_' + file))
                self.assertEqual(len(df_expected), len(df))
                self.assertEqual(len(df_expected.columns), len(df.columns))
                for i in range(len(df_expected)):
                    for j in range(len(df_expected.columns)):
                        col = df_expected.columns[j]
                        if pd.isnull(df_expected.loc[i, col]):
                            self.assertTrue(pd.isnull(df.loc[i, col]))
                        else:
                            self.assertEqual(df_expected.loc[i, col], df.loc[i, col], msg=f'i={i}; j={j}')

    # TODO :
    # - 2 cases the same day and village
