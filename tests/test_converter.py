import unittest
from unittest.mock import MagicMock, patch
import sys
import os
from pathlib import Path

# Mock win32com.client before importing converter
sys.modules["win32com"] = MagicMock()
sys.modules["win32com.client"] = MagicMock()

# Now we can import the module to be tested
# We need to add the parent directory to sys.path to import converter
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import converter

class TestConverter(unittest.TestCase):
    def setUp(self):
        # Use the mock object that converter module is actually using
        self.mock_dispatch = converter.win32com.client.Dispatch
        self.mock_app = MagicMock()
        self.mock_dispatch.return_value = self.mock_app
        
        # Reset mocks to clear history from previous tests
        self.mock_dispatch.reset_mock()
        self.mock_app.reset_mock()

    @patch("converter.Path")
    @patch("os.path.exists")
    def test_convert_ppt_to_pdf(self, mock_exists, mock_path_cls):
        # Setup mocks
        mock_path_instance = mock_path_cls.return_value
        
        mock_ppt_file = MagicMock()
        mock_ppt_file.name = "test.pptx"
        mock_ppt_file.resolve.return_value = "/abs/path/to/test.pptx"
        # Handle with_suffix chain
        mock_pdf_path = MagicMock()
        mock_pdf_path.name = "test.pdf"
        mock_pdf_path.resolve.return_value = "/abs/path/to/test.pdf"
        mock_ppt_file.with_suffix.return_value = mock_pdf_path
        
        # glob side effect: .pptx, .pptm, .ppt
        mock_path_instance.glob.side_effect = [[mock_ppt_file], [], []]
        
        mock_exists.return_value = False # PDF does not exist, so proceed
        
        # Mock Presentation object
        mock_presentation = MagicMock()
        self.mock_app.Presentations.Open.return_value = mock_presentation
        
        # Run function
        converter.convert_ppt_to_pdf("dummy_folder")
        
        # Verify interactions
        self.mock_dispatch.assert_called_with("PowerPoint.Application")
        self.mock_app.Presentations.Open.assert_called()
        mock_presentation.SaveAs.assert_called()
        mock_presentation.Close.assert_called()
        self.mock_app.Quit.assert_called()

    @patch("converter.Path")
    @patch("os.path.exists")
    def test_convert_excel_to_pdf(self, mock_exists, mock_path_cls):
        # Setup mocks
        mock_path_instance = mock_path_cls.return_value
        
        mock_excel_file = MagicMock()
        mock_excel_file.name = "test.xlsx"
        mock_excel_file.resolve.return_value = "/abs/path/to/test.xlsx"
        
        mock_pdf_path = MagicMock()
        mock_pdf_path.name = "test.pdf"
        mock_pdf_path.resolve.return_value = "/abs/path/to/test.pdf"
        mock_excel_file.with_suffix.return_value = mock_pdf_path

        # glob side effect: .xlsx, .xlsm, .xls
        mock_path_instance.glob.side_effect = [[mock_excel_file], [], []]
        
        mock_exists.return_value = False
        
        # Mock Workbook object
        mock_workbook = MagicMock()
        self.mock_app.Workbooks.Open.return_value = mock_workbook
        
        # Setup Worksheets mock to be both iterable and have attributes
        mock_worksheets = MagicMock()
        mock_workbook.Worksheets = mock_worksheets
        
        # Create individual sheet mocks
        mock_ws1 = MagicMock()
        mock_ws1.PageSetup.PrintArea = "A1:B10" # Has print area
        
        mock_ws2 = MagicMock()
        mock_ws2.PageSetup.PrintArea = None # No print area
        
        # When iterated, yield the sheets
        mock_worksheets.__iter__.return_value = iter([mock_ws1, mock_ws2])
        
        # Run function
        converter.convert_excel_to_pdf("dummy_folder")
        
        # Verify interactions
        self.mock_dispatch.assert_called_with("Excel.Application")
        self.mock_app.Workbooks.Open.assert_called()
        
        # Verify Select was called on the Worksheets collection object
        mock_worksheets.Select.assert_called()
        
        mock_workbook.ActiveSheet.ExportAsFixedFormat.assert_called()
        mock_workbook.Close.assert_called()
        self.mock_app.Quit.assert_called()

        # Verify PageSetup logic for ws2 (no print area)
        self.assertEqual(mock_ws2.PageSetup.Zoom, False)
        self.assertEqual(mock_ws2.PageSetup.FitToPagesWide, 1)
        self.assertEqual(mock_ws2.PageSetup.FitToPagesTall, False)

    @patch("converter.Path")
    @patch("os.path.exists")
    def test_convert_word_to_pdf(self, mock_exists, mock_path_cls):
        # Setup mocks
        mock_path_instance = mock_path_cls.return_value
        
        mock_word_file = MagicMock()
        mock_word_file.name = "test.docx"
        mock_word_file.resolve.return_value = "/abs/path/to/test.docx"
        
        mock_pdf_path = MagicMock()
        mock_pdf_path.name = "test.pdf"
        mock_pdf_path.resolve.return_value = "/abs/path/to/test.pdf"
        mock_word_file.with_suffix.return_value = mock_pdf_path

        # glob side effect: .docx, .docm, .doc
        mock_path_instance.glob.side_effect = [[mock_word_file], [], []]
        
        mock_exists.return_value = False
        
        # Mock Document object
        mock_document = MagicMock()
        self.mock_app.Documents.Open.return_value = mock_document
        
        # Run function
        converter.convert_word_to_pdf("dummy_folder")
        
        # Verify interactions
        self.mock_dispatch.assert_called_with("Word.Application")
        self.mock_app.Documents.Open.assert_called()
        # 17 = wdFormatPDF
        mock_document.SaveAs2.assert_called_with(mock_pdf_path.resolve.return_value, FileFormat=17)
        mock_document.Close.assert_called()
        self.mock_app.Quit.assert_called()

    @patch("converter.convert_word_to_pdf")
    @patch("converter.convert_ppt_to_pdf")
    @patch("converter.convert_excel_to_pdf")
    @patch("argparse.ArgumentParser.parse_args")
    @patch("converter.Path")
    @patch("converter.load_dotenv")
    def test_main_basic(self, mock_load_dotenv, mock_path_cls, mock_parse_args, mock_ppt, mock_excel, mock_word):
        # Setup mocks
        mock_args = MagicMock()
        mock_args.folder = "dummy_folder"
        mock_args.output = None
        mock_parse_args.return_value = mock_args
        
        mock_path_instance = mock_path_cls.return_value
        mock_path_instance.exists.return_value = True
        mock_path_instance.resolve.return_value = "/abs/path/to/dummy_folder"
        
        # Run main
        converter.main()
        
        # Verify calls
        mock_load_dotenv.assert_called_once()
        mock_ppt.assert_called()
        mock_excel.assert_called()
        mock_word.assert_called()

    @patch("converter.convert_ppt_to_pdf")
    @patch("converter.convert_excel_to_pdf")
    @patch("argparse.ArgumentParser.parse_args")
    @patch("converter.Path")
    @patch("converter.load_dotenv")
    def test_main_use_env_vars(self, mock_load_dotenv, mock_path_cls, mock_parse_args, mock_ppt, mock_excel):
        # Case: Argument is None, Env Var is Set
        mock_args = MagicMock()
        mock_args.folder = None
        mock_args.output = None
        mock_parse_args.return_value = mock_args
        
        mock_path_instance = mock_path_cls.return_value
        mock_path_instance.exists.return_value = True
        mock_path_instance.resolve.return_value = "/env/path"

        with patch.dict(os.environ, {"INPUT_FOLDER": "/env/path", "OUTPUT_FOLDER": "/env/out"}, clear=True):
            converter.main()

        # Path("/env/path") should be called
        mock_path_cls.assert_any_call("/env/path")
        # Output path from env
        mock_path_cls.assert_any_call("/env/out")
        
        mock_ppt.assert_called()

    @patch("converter.convert_ppt_to_pdf")
    @patch("converter.convert_excel_to_pdf")
    @patch("argparse.ArgumentParser.parse_args")
    @patch("converter.Path")
    @patch("converter.load_dotenv")
    def test_main_priority(self, mock_load_dotenv, mock_path_cls, mock_parse_args, mock_ppt, mock_excel):
        # Case: Argument is Set, Env Var is Set -> Argument wins
        mock_args = MagicMock()
        mock_args.folder = "/arg/path"
        mock_args.output = "/arg/out"
        mock_parse_args.return_value = mock_args
        
        mock_path_instance = mock_path_cls.return_value
        mock_path_instance.exists.return_value = True
        
        with patch.dict(os.environ, {"INPUT_FOLDER": "/env/path", "OUTPUT_FOLDER": "/env/out"}, clear=True):
            converter.main()

        # Path("/arg/path") should be called
        mock_path_cls.assert_any_call("/arg/path")
        # Output path from arg
        mock_path_cls.assert_any_call("/arg/out")

    @patch("converter.convert_ppt_to_pdf")
    @patch("converter.convert_excel_to_pdf")
    @patch("argparse.ArgumentParser.parse_args")
    @patch("converter.Path")
    @patch("converter.load_dotenv")
    def test_main_missing_config(self, mock_load_dotenv, mock_path_cls, mock_parse_args, mock_ppt, mock_excel):
        # Case: No Arg, No Env -> Exit
        mock_args = MagicMock()
        mock_args.folder = None
        mock_args.output = None
        mock_parse_args.return_value = mock_args

        with patch.dict(os.environ, {}, clear=True):
            with self.assertRaises(SystemExit) as cm:
                converter.main()
            self.assertEqual(cm.exception.code, 1)

    def test_no_files_found(self):
        # Setup mocks
        with patch("converter.Path") as mock_path_cls:
            mock_path_instance = mock_path_cls.return_value
            mock_path_instance.glob.side_effect = [[], [], [], [], [], [], [], [], []] # All empty

            # Run functions
            converter.convert_ppt_to_pdf("dummy")
            converter.convert_excel_to_pdf("dummy")
            converter.convert_word_to_pdf("dummy")

            # Dispatch should NOT be called
            self.mock_dispatch.assert_not_called()

    def test_app_launch_failure(self):
        # Setup mocks
        with patch("converter.Path") as mock_path_cls:
            mock_path_instance = mock_path_cls.return_value
            # Return dummy files to trigger dispatch
            mock_file = MagicMock()
            mock_path_instance.glob.return_value = [mock_file]
            
            # Dispatch raises exception
            self.mock_dispatch.side_effect = Exception("Launch failed")

            # Run functions (should not crash)
            converter.convert_ppt_to_pdf("dummy")
            converter.convert_excel_to_pdf("dummy")
            converter.convert_word_to_pdf("dummy")
            
            # Reset side effect for other tests
            self.mock_dispatch.side_effect = None

    @patch("converter.Path")
    @patch("os.path.exists")
    def test_pdf_already_exists(self, mock_exists, mock_path_cls):
        # Setup mocks
        mock_path_instance = mock_path_cls.return_value
        mock_file = MagicMock()
        mock_file.name = "test.pptx"
        mock_path_instance.glob.side_effect = [[mock_file], [], []] # PPT only
        
        mock_exists.return_value = True # PDF exists

        converter.convert_ppt_to_pdf("dummy")

        # Should print skip message and NOT open presentation
        self.mock_app.Presentations.Open.assert_not_called()

    @patch("converter.Path")
    @patch("os.path.exists")
    def test_conversion_exception(self, mock_exists, mock_path_cls):
        # Setup mocks
        mock_path_instance = mock_path_cls.return_value
        mock_file = MagicMock()
        mock_file.name = "test.pptx"
        mock_path_instance.glob.side_effect = [[mock_file], [], []]
        mock_exists.return_value = False

        # Open raises exception
        self.mock_app.Presentations.Open.side_effect = Exception("Open failed")

        converter.convert_ppt_to_pdf("dummy")

        # Should handle exception and Quit app
        self.mock_app.Quit.assert_called()

    @patch("converter.Path")
    @patch("os.path.exists")
    def test_output_folder_logic(self, mock_exists, mock_path_cls):
        # Setup mocks
        mock_path_instance = mock_path_cls.return_value
        mock_file = MagicMock()
        mock_file.name = "test.pptx"
        mock_file.with_suffix.return_value.name = "test.pdf" # correctly handle with_suffix().name
        mock_path_instance.glob.side_effect = [[mock_file], [], []]
        mock_exists.return_value = False
        
        output_folder = MagicMock()
        # output_folder / filename logic
        expected_pdf_path = MagicMock()
        expected_pdf_path.resolve.return_value = "/out/test.pdf"
        output_folder.__truediv__.return_value = expected_pdf_path

        converter.convert_ppt_to_pdf("dummy", output_folder)

        # Check if SaveAs called with expected path
        # We need to drill down to the mock presentation created by Open
        mock_presentation = self.mock_app.Presentations.Open.return_value
        mock_presentation.SaveAs.assert_called_with("/out/test.pdf", 32)

if __name__ == "__main__":
    unittest.main()