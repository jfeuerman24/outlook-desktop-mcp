"""Unit tests for the pure helpers in utils.addresses. No Outlook needed."""
import unittest

from outlook_desktop_mcp.utils.addresses import classify_direction

OWN = {"jfeuerman@mba2026.hbs.edu", "Jake@heartland-gp.com"}


class ClassifyDirectionTests(unittest.TestCase):
    def test_own_address_means_sent_whatever_the_case(self):
        self.assertEqual(classify_direction("JFeuerman@mba2026.hbs.edu", "Feuerman, Jake", "megray", OWN), "sent")
        self.assertEqual(classify_direction("jake@heartland-gp.com", "", "earlgrayandsons.com", OWN), "sent")

    def test_person_as_sender_means_received(self):
        self.assertEqual(classify_direction("megray@earlgrayandsons.com", "Mike Gray", "earlgrayandsons.com", OWN), "received")
        self.assertEqual(classify_direction("/o=exchangelabs/cn=abc", "Lou Castelli", "castelli", OWN), "received")

    def test_someone_else_writing_to_the_person_means_copied(self):
        self.assertEqual(classify_direction("cphilbin@mba2026.hbs.edu", "Philbin, Chris", "earlgrayandsons.com", OWN), "copied")

    def test_blank_inputs_do_not_raise(self):
        self.assertEqual(classify_direction("", "", "x", set()), "copied")
        self.assertEqual(classify_direction(None, None, None, []), "copied")


if __name__ == "__main__":
    unittest.main()
