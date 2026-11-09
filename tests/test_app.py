import io
import unittest

from pptx import Presentation

import app


class AppTests(unittest.TestCase):
    def make_pptx(self):
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[5])
        slide.shapes.title.text = "Tujuan"
        output = io.BytesIO()
        presentation.save(output)
        return output.getvalue()

    def test_oversized_upload_returns_clear_413_instead_of_500(self):
        original_limit = app.app.config["MAX_CONTENT_LENGTH"]
        app.app.config["MAX_CONTENT_LENGTH"] = 32
        try:
            response = app.app.test_client().post(
                "/api/docx/start",
                data={"presentation": (io.BytesIO(b"x" * 100), "large.pptx")},
                content_type="multipart/form-data",
            )
        finally:
            app.app.config["MAX_CONTENT_LENGTH"] = original_limit

        self.assertEqual(response.status_code, 413)
        self.assertIn("melebihi batas", response.json["error"].lower())


if __name__ == "__main__":
    unittest.main()
