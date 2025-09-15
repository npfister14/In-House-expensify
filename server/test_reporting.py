import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from reporting import buildTables, formatCHF, percent, summarizeForAI


class DummyChoice:
    def __init__(self, content: str):
        self.message = type('Message', (), {'content': content})()


class DummyResponse:
    def __init__(self, content: str):
        self.choices = [DummyChoice(content)]


class DummyCompletions:
    def __init__(self, content: str):
        self._content = content

    def create(self, **kwargs):  # pragma: no cover - simple stub
        return DummyResponse(self._content)


class DummyClient:
    def __init__(self, content: str):
        self.chat = type('Chat', (), {'completions': DummyCompletions(content)})()


class ReportingHelpersTest(unittest.TestCase):
    def test_format_chf(self):
        self.assertEqual(formatCHF(1234.5), "CHF 1'234.50")
        self.assertEqual(formatCHF(None), "CHF 0.00")

    def test_percent(self):
        self.assertEqual(percent(50, 200), '25.0%')
        self.assertEqual(percent(0, 0), '0.0%')

    def test_build_tables_converts_to_chf(self):
        report = {
            'period': '2024-05',
            'rows': [
                {
                    'date': '2024-05-01',
                    'payer': 'Alice',
                    'category': 'Meals',
                    'paymentMethod': 'Company card',
                    'gross': 100,
                    'net': 92,
                    'vat': 8,
                    'currency': 'USD',
                    'status': 'Done',
                },
                {
                    'date': '2024-05-02',
                    'payer': 'Bob',
                    'category': 'Travels',
                    'paymentMethod': 'Personal',
                    'gross': 200,
                    'net': 180,
                    'vat': 20,
                    'currency': 'EUR',
                    'status': 'Done',
                },
                {
                    'date': '2024-05-03',
                    'payer': 'Cara',
                    'category': 'Meals',
                    'paymentMethod': 'Cash',
                    'gross': 50,
                    'net': 45,
                    'vat': 5,
                    'currency': 'CHF',
                    'status': 'In-Progress',
                },
                {
                    'date': '2024-05-04',
                    'payer': 'Dana',
                    'category': 'Supplies',
                    'paymentMethod': 'Company card',
                    'gross': 80,
                    'net': 74,
                    'vat': 6,
                    'currency': 'USD',
                    'status': 'Under Review',
                },
            ],
            'fxRatesCHF': {'USD': 0.9, 'EUR': 0.96, 'CHF': 1.0},
            'fxPolicy': 'Test policy',
        }

        tables = buildTables(report)

        self.assertEqual(tables['totals']['grossCHF'], 282.0)
        self.assertEqual(tables['totals']['netCHF'], 255.6)
        self.assertEqual(tables['totals']['vatCHF'], 26.4)
        self.assertEqual(tables['totals']['companyCardSpentCHF'], 90.0)
        self.assertEqual(tables['totals']['reimbursementsOwedCHF'], 192.0)
        self.assertEqual(tables['pending']['inProgress']['amount'], 50.0)
        self.assertEqual(tables['pending']['underReview']['amount'], 72.0)
        self.assertEqual(len(tables['rowsCHF']), 4)
        self.assertEqual(tables['rowsCHF'][0]['amountCHF'], 90.0)
        self.assertEqual(tables['rowsCHF'][1]['amountCHF'], 192.0)
        self.assertEqual(tables['topCategory'], ('Travels', 192.0))
        self.assertEqual(tables['reimbursements'], [('Bob', 192.0)])

    def test_summarize_for_ai_uses_client(self):
        metrics = {
            'period': '2024-05',
            'totals': {'grossCHF': 100.0, 'companyCardSpentCHF': 60.0},
            'topCategory': ('Meals', 70.0),
            'topOwed': [('Bob', 30.0)],
            'pending': {'inProgress': {'count': 1, 'amount': 10.0}, 'underReview': {'count': 0, 'amount': 0.0}},
        }
        client = DummyClient('Summary text here.')
        summary = summarizeForAI(metrics, client=client)
        self.assertEqual(summary, 'Summary text here.')


if __name__ == '__main__':
    unittest.main()
