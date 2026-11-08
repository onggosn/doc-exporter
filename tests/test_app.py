import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

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

    def test_generate_endpoint_extracts_ppt_and_returns_generated_document(self):
        with tempfile.TemporaryDirectory() as upload_dir:
            with patch.object(app, "UPLOAD_DIR", Path(upload_dir)), patch.object(
                app, "BedrockDocumentGenerator"
            ) as generator_class:
                generator_class.return_value.generate.return_value = {
                    "filename": "sample.pptx",
                    "chapters": [
                        {"title": "Bab 1", "html": "<h2>Bab 1</h2>", "slide_refs": ["slide-1"]}
                    ],
                    "html": "<h2>Bab 1</h2>",
                }
                response = app.app.test_client().post(
                    "/api/generate",
                    data={
                        "presentation": (io.BytesIO(self.make_pptx()), "sample.pptx"),
                        "chapters": "Bab 1",
                    },
                    content_type="multipart/form-data",
                )

            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.json["filename"], "sample.pptx")
            self.assertEqual(list(Path(upload_dir).iterdir()), [])

    def test_oversized_upload_returns_clear_413_instead_of_500(self):
        original_limit = app.app.config["MAX_CONTENT_LENGTH"]
        app.app.config["MAX_CONTENT_LENGTH"] = 32
        try:
            response = app.app.test_client().post(
                "/api/generate",
                data={"presentation": (io.BytesIO(b"x" * 100), "large.pptx")},
                content_type="multipart/form-data",
            )
        finally:
            app.app.config["MAX_CONTENT_LENGTH"] = original_limit

        self.assertEqual(response.status_code, 413)
        self.assertIn("melebihi batas", response.json["error"].lower())


if __name__ == "__main__":
    unittest.main()
