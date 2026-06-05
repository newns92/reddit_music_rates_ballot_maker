'''
To run:
pip install pytest
pytest test_add_ballot_to_excel.py -v
'''

# test_add_ballot_to_excel.py
import pytest
from unittest.mock import MagicMock, patch, call
from add_ballot_to_excel import find_cells_by_value, create_function_strings, insert_stats_cells


# ── Helpers ────────────────────────────────────────────────────────────────────

def make_cell(value, row, column=1):
    """Creates a mock openpyxl cell object."""
    cell = MagicMock()
    cell.value = value
    cell.row = row
    cell.column = column
    cell.coordinate = f"{chr(64 + column)}{row}"
    return cell


def make_final_cells(bonus_row=54, end_row=70, album_rows=(2, 14, 27, 41)):
    """Builds a final_cells dict using mock cells, mirroring the real structure."""
    cells = {
        'bonus': make_cell('BONUS TRACKS', bonus_row),
        'end':   make_cell('END', end_row),
    }
    for i, row in enumerate(album_rows, start=1):
        cells[f'album_{i}'] = make_cell(f'Album: Test {i}', row)
    return cells


# ── create_function_strings ────────────────────────────────────────────────────

class TestCreateFunctionStrings:

    def test_album_ranges_skip_header_row(self):
        """Range start should be album_row + 1 (skipping the 'Album:' header)."""
        final_cells = make_final_cells(album_rows=(2,), bonus_row=10, end_row=15)
        totals, _ = create_function_strings(final_cells)
        assert 'B3' in totals['Average']   # starts at row 3, not 2

    def test_album_range_ends_before_next_album(self):
        """First album range should end one row above the second album."""
        final_cells = make_final_cells(album_rows=(2, 14), bonus_row=25, end_row=30)
        totals, _ = create_function_strings(final_cells)
        assert 'B3:B13' in totals['Average']

    def test_last_album_range_ends_before_bonus(self):
        """Last album range should end one row above BONUS TRACKS."""
        final_cells = make_final_cells(album_rows=(2, 14), bonus_row=25, end_row=30)
        totals, _ = create_function_strings(final_cells)
        assert 'B15:B24' in totals['Average']

    def test_all_totals_functions_present(self):
        """All six stat keys should be returned."""
        final_cells = make_final_cells()
        totals, _ = create_function_strings(final_cells)
        assert set(totals.keys()) == {'Average', 'Median', 'Tens', 'SubFives', 'Mode', 'StdDev'}

    def test_no_at_sign_in_formulas(self):
        """No formula should contain '@' (implicit intersection operator)."""
        final_cells = make_final_cells()
        totals, album_avgs = create_function_strings(final_cells)
        for formula in list(totals.values()) + list(album_avgs.values()):
            assert '@' not in formula, f"Unexpected '@' in: {formula}"

    def test_no_curly_braces_in_formulas(self):
        """No formula should contain '{}' (array formula notation)."""
        final_cells = make_final_cells()
        totals, album_avgs = create_function_strings(final_cells)
        for formula in list(totals.values()) + list(album_avgs.values()):
            assert '{' not in formula and '}' not in formula

    def test_xlfn_prefix_present(self):
        """All formulas should use the _xlfn. prefix."""
        final_cells = make_final_cells()
        totals, album_avgs = create_function_strings(final_cells)
        for formula in list(totals.values()) + list(album_avgs.values()):
            assert formula.startswith('=_xlfn.'), f"Missing _xlfn. prefix in: {formula}"

    def test_album_avg_functions_keyed_by_cell(self):
        """album_avg_functions keys should be cell addresses like 'B2', 'B14'."""
        final_cells = make_final_cells(album_rows=(2, 14), bonus_row=25, end_row=30)
        _, album_avgs = create_function_strings(final_cells)
        assert 'B2' in album_avgs
        assert 'B14' in album_avgs

    def test_bonus_avg_function_included(self):
        """album_avg_functions should include an entry for the BONUS TRACKS row."""
        final_cells = make_final_cells(bonus_row=54, end_row=70)
        _, album_avgs = create_function_strings(final_cells)
        assert 'B54' in album_avgs
        assert 'B55:B69' in album_avgs['B54']

    def test_four_albums_produces_four_ranges(self):
        """Four albums should produce four ranges joined in the formula."""
        final_cells = make_final_cells(album_rows=(2, 14, 27, 41), bonus_row=54, end_row=70)
        totals, _ = create_function_strings(final_cells)
        # Four ranges = three commas
        assert totals['Average'].count(',') == 3

    def test_album_order_is_by_row(self):
        """Albums inserted out of order in the dict should still produce correct ranges."""
        cells = {
            'bonus':   make_cell('BONUS TRACKS', 25),
            'end':     make_cell('END', 30),
            'album_2': make_cell('Album: B', 14),   # intentionally out of order
            'album_1': make_cell('Album: A', 2),
        }
        totals, _ = create_function_strings(cells)
        avg = totals['Average']
        # B3 (album_1 start) should appear before B15 (album_2 start)
        assert avg.index('B3') < avg.index('B15')


# ── find_cells_by_value ────────────────────────────────────────────────────────

class TestFindCellsByValue:

    def _make_ws(self, rows):
        """Builds a mock worksheet from a list of (value, row) tuples per row."""
        mock_rows = []
        for row_cells in rows:
            mock_row = [make_cell(v, r) for v, r in row_cells]
            mock_rows.append(mock_row)
        ws = MagicMock()
        ws.iter_rows.return_value = mock_rows
        return ws

    @patch('add_ballot_to_excel.load_workbook')
    def test_finds_bonus_and_end(self, mock_load):
        ws = self._make_ws([
            [('Album: Foo', 1)],
            [('track', 2)],
            [('BONUS TRACKS', 3)],
            [('END', 4)],
        ])
        mock_load.return_value.__getitem__.return_value = ws
        mock_load.return_value.close = MagicMock()

        result = find_cells_by_value('f.xlsx', 'Sheet1', 'BONUS TRACKS', 'END')

        assert result['bonus'].value == 'BONUS TRACKS'
        assert result['end'].value == 'END'

    @patch('add_ballot_to_excel.load_workbook')
    def test_finds_all_albums(self, mock_load):
        ws = self._make_ws([
            [('Album: A', 1)],
            [('Album: B', 5)],
            [('BONUS TRACKS', 9)],
            [('END', 12)],
        ])
        mock_load.return_value.__getitem__.return_value = ws
        mock_load.return_value.close = MagicMock()

        result = find_cells_by_value('f.xlsx', 'Sheet1', 'BONUS TRACKS', 'END')

        assert 'album_1' in result
        assert 'album_2' in result
        assert result['album_1'].row == 1
        assert result['album_2'].row == 5
