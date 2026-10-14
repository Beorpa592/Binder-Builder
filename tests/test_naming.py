import pytest

from binder_builder.engine.codes import filename_part, letter_sequence
from binder_builder.engine.tree import slugify


@pytest.mark.parametrize("title,expected", [
    ("Open Points / Action Items", "Open_Points_Action_Items"),
    ("Index of Genres & Channels", "Index_of_Genres_and_Channels"),
    ("  North Star   MRP Summary ", "North_Star_MRP_Summary"),
    ("Top 10 List", "Top_10_List"),
    ("Q&A", "QandA"),
    ("Café (draft) #2", "Caf_draft_2"),
    ("a - b", "a_-_b"),
    ("__x__", "x"),
])
def test_filename_part(title, expected):
    assert filename_part(title) == expected


@pytest.mark.parametrize("title,expected", [
    ("Top 10 List", "top-10-list"),
    ("Index of Genres & Channels", "index-of-genres-and-channels"),
    ("Open Points / Action Items", "open-points-action-items"),
    ("!!!", "node"),
])
def test_slugify(title, expected):
    assert slugify(title) == expected


def test_letters_skip_nothing_and_roll_over():
    gen = letter_sequence()
    letters = [next(gen) for _ in range(30)]
    assert letters[:3] == ["A", "B", "C"]
    assert letters[25] == "Z"
    assert letters[26:28] == ["AA", "AB"]
