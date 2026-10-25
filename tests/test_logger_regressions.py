from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from qt_app.services.logger import AppLogger


class LoggerRegressionTests(unittest.TestCase):
    def test_close_releases_log_for_directory_cleanup_and_does_not_reopen(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "output.log"
            logger = AppLogger(path)
            self.addCleanup(logger.close)
            logger.log("before close")
            stream = logger._handler.stream
            logger.close()
            logger.close()
            self.assertTrue(stream.closed)
            path.unlink()  # Windows requires the file handle to be closed.
            logger.log("after close")
            self.assertFalse(path.exists())

    def test_rotation_preserves_each_record_once(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "output.log"
            logger = AppLogger(path)
            self.addCleanup(logger.close)
            logger._handler.maxBytes = 70
            for index in range(4):
                logger.log(f"record-{index}")
            logger.close()
            files = list(Path(directory).glob("output.log*"))
            text = "".join(file.read_text(encoding="utf-8") for file in files)
            self.assertGreater(len(files), 1)
            for index in range(4):
                self.assertEqual(text.count(f"record-{index}"), 1)


if __name__ == "__main__":
    unittest.main()
