"""Simple Python tests for a media cabinet inventory."""

import unittest


class MediaCabinet:
    def __init__(self):
        self._items = {}

    def add_item(self, title: str, count: int = 1) -> None:
        if count < 1:
            raise ValueError("count must be >= 1")
        self._items[title] = self._items.get(title, 0) + count

    def remove_item(self, title: str, count: int = 1) -> None:
        if title not in self._items:
            raise KeyError(f"{title} does not exist")
        if count < 1:
            raise ValueError("count must be >= 1")
        if self._items[title] < count:
            raise ValueError("not enough items to remove")

        self._items[title] -= count
        if self._items[title] == 0:
            del self._items[title]

    def count(self, title: str) -> int:
        return self._items.get(title, 0)


class TestMediaCabinet(unittest.TestCase):
    def test_add_item_increases_count(self):
        cabinet = MediaCabinet()
        cabinet.add_item("Interstellar")
        cabinet.add_item("Interstellar", 2)
        self.assertEqual(cabinet.count("Interstellar"), 3)

    def test_remove_item_decreases_count(self):
        cabinet = MediaCabinet()
        cabinet.add_item("Inception", 3)
        cabinet.remove_item("Inception", 2)
        self.assertEqual(cabinet.count("Inception"), 1)

    def test_remove_all_items_deletes_entry(self):
        cabinet = MediaCabinet()
        cabinet.add_item("Dune", 1)
        cabinet.remove_item("Dune", 1)
        self.assertEqual(cabinet.count("Dune"), 0)

    def test_invalid_count_raises_error(self):
        cabinet = MediaCabinet()
        with self.assertRaises(ValueError):
            cabinet.add_item("Arrival", 0)


if __name__ == "__main__":
    unittest.main()
