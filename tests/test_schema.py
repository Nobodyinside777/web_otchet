from app.schema import COLUMNS, col_index, col_letter


def test_columns_go_from_b_to_aa_without_gaps():
    assert [col_index(c.letter) for c in COLUMNS] == list(range(col_index("B"), col_index("AA") + 1))


def test_codes_and_labels_are_unique():
    assert len({c.code for c in COLUMNS}) == len(COLUMNS)
    assert len({c.label for c in COLUMNS}) == len(COLUMNS)


def test_letter_conversion():
    assert col_index("A") == 0
    assert col_index("Z") == 25
    assert col_index("AA") == 26
    for i in range(80):
        assert col_index(col_letter(i)) == i
