import sys
import unittest
from pathlib import Path
from unittest import mock

import main


class SettingsFilePathTests(unittest.TestCase):
    def test_sits_next_to_the_source_file_when_running_from_source(self):
        with mock.patch.object(sys, "frozen", False, create=True):
            path = Path(main.settings_file_path())

        self.assertEqual(path.parent, Path(main.__file__).resolve().parent)
        self.assertEqual(path.name, "ScanSweep.ini")

    def test_sits_next_to_the_executable_when_frozen(self):
        fake_exe = Path(r"D:\Portable\ScanSweep\ScanSweep-2.0.exe")
        with mock.patch.object(sys, "frozen", True, create=True), \
             mock.patch.object(sys, "executable", str(fake_exe)):
            path = Path(main.settings_file_path())

        self.assertEqual(path.parent, fake_exe.parent)
        self.assertEqual(path.name, "ScanSweep.ini")


if __name__ == "__main__":
    unittest.main()
