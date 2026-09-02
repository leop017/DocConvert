import unittest


class TestPackageInit(unittest.TestCase):

    def test_import_docconvert_app_lazy(self):
        import docconvert
        app_cls = docconvert.DocConvertApp
        self.assertEqual(app_cls.__name__, "DocConvertApp")

    def test_attribute_error_for_unknown_attr(self):
        import docconvert
        with self.assertRaises(AttributeError):
            _ = docconvert.NonExistentSymbol

    def test_all_contains_docconvert_app(self):
        import docconvert
        self.assertIn("DocConvertApp", docconvert.__all__)


if __name__ == "__main__":
    unittest.main()
