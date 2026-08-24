import io
import logging
import unittest

from docconvert.logger import LOGGER_NAME, get_logger, setup_logging


class TestSetupLogging(unittest.TestCase):
    """Regression: a second setup_logging call (e.g. main.py INFO first,
    then main_cli --verbose DEBUG) must raise the already-registered
    handler's level, otherwise DEBUG messages are silently filtered out.
    """

    def setUp(self):
        self.logger = get_logger()
        self._orig_handlers = list(self.logger.handlers)
        self._orig_level = self.logger.level
        for h in self.logger.handlers:
            self.logger.removeHandler(h)

    def tearDown(self):
        for h in self.logger.handlers:
            self.logger.removeHandler(h)
        for h in self._orig_handlers:
            self.logger.addHandler(h)
        self.logger.setLevel(self._orig_level)

    def _attach_buffer(self):
        buf = io.StringIO()
        handler = logging.StreamHandler(buf)
        handler.setFormatter(logging.Formatter('%(message)s'))
        self.logger.addHandler(handler)
        return buf

    def test_level_syncs_to_existing_handler(self):
        setup_logging(level='INFO')          # first call creates the handler
        self._attach_buffer()                # our own capture handler
        setup_logging(level='DEBUG')         # second call (--verbose path)
        self.logger.debug('SENTINEL_DEBUG')
        self.assertIn('SENTINEL_DEBUG', self.logger.handlers[-1].stream.getvalue(),
                      "DEBUG must pass an existing handler after re-setup")

    def test_second_call_same_level_still_logs(self):
        setup_logging(level='DEBUG')
        buf = self._attach_buffer()
        setup_logging(level='DEBUG')
        self.logger.info('SENTINEL_INFO')
        self.assertIn('SENTINEL_INFO', buf.getvalue())

    def test_returns_logger(self):
        self.assertIs(setup_logging(), get_logger())


if __name__ == '__main__':
    unittest.main()
