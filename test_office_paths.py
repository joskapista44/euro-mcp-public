#!/usr/bin/env python3
"""Tests for the office path guard.

The guard decides whether a tool may touch a file at all, so the refusals are the
product. Two of them are the reason it exists: a symlink that points out of an
allowed directory, and a path that merely LOOKS like it is inside one
("/tmpfoo" vs "/tmp"). Both would pass a naive prefix check.

Run:  python3 -m unittest test_office_paths
"""

import os
import tempfile
import unittest

from office_paths import DEFAULT_ROOTS, PathNotAllowed, allowed_roots, check_path


class AllowedRootsTests(unittest.TestCase):
    def test_default_is_tmp_where_the_fleet_downloads_files(self):
        self.assertEqual(allowed_roots({}), [os.path.realpath(DEFAULT_ROOTS)])

    def test_colon_and_comma_are_both_accepted(self):
        env = {"OFFICE_ALLOWED_ROOTS": "/tmp:/var/tmp"}
        self.assertEqual(allowed_roots(env), [os.path.realpath("/tmp"), os.path.realpath("/var/tmp")])
        env = {"OFFICE_ALLOWED_ROOTS": "/tmp,/var/tmp"}
        self.assertEqual(allowed_roots(env), [os.path.realpath("/tmp"), os.path.realpath("/var/tmp")])

    def test_empty_entries_do_not_become_the_filesystem_root(self):
        # "/tmp::" must not turn into an allowed "" that matches everything.
        self.assertEqual(allowed_roots({"OFFICE_ALLOWED_ROOTS": "/tmp::"}), [os.path.realpath("/tmp")])


class CheckPathTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="office-guard-")
        self.env = {"OFFICE_ALLOWED_ROOTS": self.tmp}

    def test_a_file_inside_an_allowed_root_passes(self):
        target = os.path.join(self.tmp, "book.xlsx")
        open(target, "w").close()
        self.assertEqual(check_path(target, self.env), os.path.realpath(target))

    def test_a_file_that_does_not_exist_yet_still_passes_inside_the_root(self):
        # Writing a NEW file in an allowed directory is legitimate.
        self.assertTrue(check_path(os.path.join(self.tmp, "new.xlsx"), self.env))

    def test_a_path_outside_every_root_is_REFUSED(self):
        # Any path outside every allowed root works here -- this one is arbitrary, not special.
        with self.assertRaises(PathNotAllowed):
            check_path("/etc/some-other-service/config.db", self.env)

    def test_a_SYMLINK_pointing_out_of_the_root_is_REFUSED(self):
        # The spelling is inside the root; where it lands is not. This is the
        # case a prefix check on the raw string would happily allow.
        outside = tempfile.mkdtemp(prefix="office-outside-")
        link = os.path.join(self.tmp, "escape")
        os.symlink(outside, link)
        with self.assertRaises(PathNotAllowed):
            check_path(os.path.join(link, "book.xlsx"), self.env)

    def test_dotdot_traversal_is_REFUSED(self):
        with self.assertRaises(PathNotAllowed):
            check_path(os.path.join(self.tmp, "..", "..", "etc", "passwd"), self.env)

    def test_a_sibling_whose_NAME_starts_with_the_root_is_REFUSED(self):
        # "/tmpfoo" is not inside "/tmp": the boundary is a separator, not a prefix.
        sibling = self.tmp + "foo"
        os.makedirs(sibling, exist_ok=True)
        try:
            with self.assertRaises(PathNotAllowed):
                check_path(os.path.join(sibling, "book.xlsx"), self.env)
        finally:
            os.rmdir(sibling)

    def test_the_root_itself_is_allowed(self):
        self.assertTrue(check_path(self.tmp, self.env))

    def test_an_empty_root_list_refuses_EVERYTHING(self):
        # Fails closed: a misconfiguration must not silently mean "allow all".
        with self.assertRaises(PathNotAllowed):
            check_path(os.path.join(self.tmp, "book.xlsx"), {"OFFICE_ALLOWED_ROOTS": ""})

    def test_the_refusal_names_the_resolved_path_and_the_roots(self):
        # An operator has to be able to act on the message without reading the source.
        try:
            check_path("/etc/passwd", self.env)
        except PathNotAllowed as exc:
            self.assertIn("/etc/passwd", str(exc))
            self.assertIn(self.tmp, str(exc))
        else:
            self.fail("expected a refusal")


if __name__ == "__main__":
    unittest.main()
