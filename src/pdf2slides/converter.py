import io
import os.path

import numpy as np
import pymupdf
from paddleocr import PaddleOCR, PPStructure
from PIL import Image
from pptx import Presentation, presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.slide import Slide
from pptx.util import Inches, Pt
from sklearn.mixture import GaussianMixture
from sklearn.model_selection import GridSearchCV


class Converter:

    def __init__(
        self,
        default_font: str | None = None,
        enable_ocr: bool = False,
        enforce_default_font: bool = False,
        image_retention_level: float = 1.0,
        lang: str = "ch",
    ) -> None:
        """
        Args:
            default_font: This is the font that will be used for OCR'ed text. Does not have any effect if the font is not already installed in the computer.
            enable_ocr: Whether the OCR will be used. Defaults to `False`.
            enforce_default_font: If `true`, use the default font even when OCR is not enabled. Otherwise, use the font detected in the original PDF file. Defaults to `False`.
            image_retention_level: Determines how likely pure-text images will be kept. Has no effect when not in OCR mode. This is internally passed to PaddleOCR as the `layout_score_threshold` parameter. Defaults to `1.0`.
            lang: The abbreviation of the language used by the PaddleOCR. Defaults to `ch`, which supports both Chinese and English. For a full list of supported languages, please refer to https://github.com/PaddlePaddle/PaddleOCR/blob/main/doc/doc_en/multi_languages_en.md.
        """

        self.default_font = default_font
        self.enable_ocr = enable_ocr
        self.enforce_default_font = enforce_default_font

        self.ocr = PaddleOCR(lang=lang, show_log=False) if self.enable_ocr else None
        self.layout_engine = (
            PPStructure(
                table=False,
                ocr=False,
                layout_score_threshold=image_retention_level,
            )
            if self.enable_ocr
            else None
        )

    def convert(self, input_file_path: str, output_file_path: str) -> None:
        """Convert a .pdf file to a .pptx file.

        Args:
            input_file_path: The path of the input `.pdf` file.
            output_file_path: The path of the generated `.pptx` file.
        """

        # Validate the arguments. Raise an error if they are invalid.
        Converter._validate_arguments(input_file_path, output_file_path)

        pptx_output = Presentation()

        with pymupdf.open(input_file_path) as pdf_document:
            # Determine slide size from the first page
            first_page = pdf_document[0]
            first_page_width = first_page.rect.width
            first_page_height = first_page.rect.height

            # Convert PDF page size from points to inches
            slide_width = first_page_width / 72.0
            slide_height = first_page_height / 72.0

            # Set the slide size as the matching size
            pptx_output.slide_width = Inches(slide_width)
            pptx_output.slide_height = Inches(slide_height)

            all_contents = [
                self._get_page_contents(pdf_page) for pdf_page in pdf_document
            ]

            # No page contains editable text. The document is likely scanned.
            scanned_document = self.enable_ocr and all(
                (not page_contents["text_blocks"]) for page_contents in all_contents
            )

            self._construct_pptx(
                pdf_document, all_contents, pptx_output, scanned_document
            )

        output_directory = os.path.dirname(output_file_path)
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)

        # Save PowerPoint presentation
        pptx_output.save(output_file_path)

    def _get_page_contents(self, pdf_page: pymupdf.Page) -> dict:
        """Get all editable contents (excluding OCR'ed text) from page."""

        page_contents = {}
        page_contents["text_blocks"] = self._get_page_text(pdf_page)
        page_contents["drawings"] = self._get_page_drawings(pdf_page)
        page_contents["images"] = self._get_page_images(pdf_page)
        page_contents["xref_smask_map"] = self._get_page_xref_smask_map(pdf_page)
        return page_contents

    def _get_page_text(self, pdf_page: pymupdf.Page) -> list[dict]:
        """Get all text blocks from the page."""

        blocks = pdf_page.get_textpage().extractDICT()["blocks"]
        text_blocks = [block for block in blocks if block["type"] == 0]
        return text_blocks

    def _get_page_drawings(self, pdf_page: pymupdf.Page) -> list[dict]:
        """Get all drawings from the page."""

        drawings = pdf_page.get_drawings()
        return drawings

    def _get_page_images(self, pdf_page: pymupdf.Page) -> list[dict]:
        """Get all images from the page."""

        images = pdf_page.get_image_info(xrefs=True)
        return images

    def _get_page_xref_smask_map(self, pdf_page: pymupdf.Page) -> dict[int, int]:
        """Store the smask of the images in a dictionary."""

        xref_smask_map = {item[0]: item[1] for item in pdf_page.get_images()}
        return xref_smask_map

    def _get_page_ocr_text(self, pdf_page: pymupdf.Page) -> list:
        """OCR the page and get the result."""

        page_pixmap = pdf_page.get_pixmap(dpi=300)
        page_bytes = page_pixmap.tobytes()
        ocr_result = self.ocr.ocr(page_bytes, cls=False)
        return ocr_result

    def _add_text_block_to_slide(self, text_block: dict, slide: Slide) -> None:
        """Add all text in the block to a slide. Output font properties will follow those found in the original file."""

        for line in text_block["lines"]:
            for span in line["spans"]:
                # Extract text content and font size from each span
                line_text = span["text"]
                line_font_size = span["size"]
                # The font size is too small, will raise an error when added to slide
                if line_font_size < 1:
                    continue
                line_font_name = span["font"]
                line_font_color = span["color"]
                line_font_is_italic = bool(span["flags"] & 2**1)
                line_font_is_bold = bool(span["flags"] & 2**4)

                # Extract position of the span
                x0, y0, x1, y1 = span["bbox"]

                # Calculate position on the slide
                # Convert from points to inches (1 point = 1/72 inch)
                text_box_left = Inches(x0 / 72.0)
                text_box_top = Inches(y0 / 72.0)
                text_box_width = Inches((x1 - x0) / 72.0)
                text_box_height = Inches((y1 - y0) / 72.0)

                # Add text box to the slide
                new_text_box = slide.shapes.add_textbox(
                    text_box_left, text_box_top, text_box_width, text_box_height
                )

                text_frame = new_text_box.text_frame
                # Centers text vertically. Necessary for placing mathematical formulae in Beamer at desired locations.
                text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                # Set the margin to 0 to visually mimic original layout
                text_frame.margin_bottom = 0
                text_frame.margin_left = 0
                text_frame.margin_right = 0
                text_frame.margin_top = 0

                p = text_frame.paragraphs[0]
                p.text = line_text
                p.font.size = Pt(line_font_size)
                p.font.name = (
                    self.default_font
                    if self.enforce_default_font and self.default_font
                    else line_font_name
                )
                # BUG: Text with gradient color will become black
                p.font.color.rgb = RGBColor.from_string(f"{line_font_color:06X}")
                p.font.italic = line_font_is_italic
                p.font.bold = line_font_is_bold
                p.alignment = PP_ALIGN.LEFT

    def _add_image_to_slide(
        self,
        pdf_doc: pymupdf.Document,
        image: dict,
        smask: int,
        bbox: pymupdf.Rect,
        slide: Slide,
    ) -> None:
        """Add the image to a slide."""

        xref = image["xref"]
        base_image = pymupdf.Pixmap(pdf_doc, xref)

        # The image itself is a smask. We ingore this image.
        if not base_image.colorspace:
            return

        # The transparency data is contained in the smask
        if smask:
            smask_pixmap = pymupdf.Pixmap(pdf_doc, smask)
            # The alpha channel is already present but contains no information
            if base_image.alpha:
                base_image.set_alpha(smask_pixmap.samples)
            # The alpha channel is not present
            else:
                base_image = pymupdf.Pixmap(base_image, smask_pixmap)

        transformation_matrix = pymupdf.Matrix(image["transform"])
        transformed_image_bytes = Converter._transform_image(
            base_image, bbox, transformation_matrix
        )

        # Calculate position on the slide based on original PDF coordinates
        left = Inches(bbox[0] / 72.0)
        top = Inches(bbox[1] / 72.0)
        width = Inches((bbox[2] - bbox[0]) / 72.0)

        # Add image to the PowerPoint slide
        slide.shapes.add_picture(
            io.BytesIO(transformed_image_bytes), left, top, width=width
        )

    def _add_drawing_to_slide(
        self,
        drawing: dict,
        bbox: pymupdf.Rect,
        pdf_page_width: float,
        pdf_page_height: float,
        slide: Slide,
    ) -> None:
        """Add all drawings to a slide.

        Reference: https://pymupdf.readthedocs.io/en/latest/recipes-drawing-and-graphics.html
        """
        # BUG: Drawings with gradient color will be lost

        # Create a temporary page with the same dimensions to serve as a drawing canvas for the shape
        # NOTE: Empirically, creating a new page for each shape is no slower than drawing each shape and then removing it on the same page
        temp_pdf = pymupdf.open()
        temp_page = temp_pdf.new_page(width=pdf_page_width, height=pdf_page_height)

        shape = temp_page.new_shape()
        for item in drawing["items"]:  # these are the draw commands
            if item[0] == "l":  # line
                shape.draw_line(item[1], item[2])
            elif item[0] == "re":  # rectangle
                shape.draw_rect(item[1])
            elif item[0] == "qu":  # quad
                shape.draw_quad(item[1])
            elif item[0] == "c":  # curve
                shape.draw_bezier(item[1], item[2], item[3], item[4])

        keys = [
            "fill",
            "color",
            "dashes",
            "even_odd",
            "closePath",
            "lineJoin",
            "width",
            "stroke_opacity",
            "fill_opacity",
        ]
        kwargs_defaults = {
            "even_odd": True,
            "closePath": True,
            "lineJoin": 0,
            "width": 1,
            "stroke_opacity": 1,
            "fill_opacity": 1,
        }
        kwargs = {
            key: (
                drawing.get(key)
                if drawing.get(key) is not None
                else kwargs_defaults.get(key)
            )
            for key in keys
        }
        kwargs["lineCap"] = (
            max(drawing.get("lineCap")) if drawing.get("lineCap") is not None else 0
        )
        shape.finish(**kwargs)
        shape.commit()

        base_image = temp_page.get_pixmap(clip=bbox, dpi=300, alpha=True)
        try:
            # The image has .png format by default
            image_bytes = base_image.tobytes()

            # Calculate position on the slide based on original PDF coordinates
            left = Inches(bbox[0] / 72.0)
            top = Inches(bbox[1] / 72.0)
            width = Inches((bbox[2] - bbox[0]) / 72.0)

            # Add image to the PowerPoint slide
            slide.shapes.add_picture(io.BytesIO(image_bytes), left, top, width=width)
        except:
            # Something went wrong. Ignore this drawing.
            pass

    def _add_ocr_to_slide(self, ocr_line: list, fontsize: float, slide: Slide) -> None:
        """Add all text obtained from ocr to a slide."""

        # Since the input pixmap has dpi=300 rather than dpi=72 by default, we need to set the scaling factor as 72/300 = 0.24
        scaling_factor = 0.24

        text = ocr_line[1][0]
        rect = ocr_line[0]

        # Get bounding box coordinates
        left, top, right, bottom = Converter._get_bbox(rect)
        left *= scaling_factor
        top *= scaling_factor
        right *= scaling_factor
        bottom *= scaling_factor

        # Calculate position and size in PowerPoint slide
        x, y = left, top
        width, height = right - left, bottom - top

        # Add text box to slide
        text_box = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(width), Pt(height))
        text_frame = text_box.text_frame
        text_frame.text = text

        # Set the margin to 0 to visually mimic original layout
        text_frame.margin_bottom = 0
        text_frame.margin_left = 0
        text_frame.margin_right = 0
        text_frame.margin_top = 0

        for paragraph in text_frame.paragraphs:
            # Set font size
            paragraph.font.size = Pt(fontsize)
            # Set the font family to the default font
            if self.default_font:
                paragraph.font.name = self.default_font

    def _construct_pptx(
        self,
        pdf_document: pymupdf.Document,
        all_contents: list[dict],
        pptx_output: presentation.Presentation,
        scanned_document: bool,
    ) -> None:
        """Construct output slides with scanned PDF file."""

        if scanned_document:
            ocr_results = [
                self._get_page_ocr_text(pdf_page) for pdf_page in pdf_document
            ]
            detected_sizes = [
                Converter._get_ocr_detected_font_size(line)
                for ocr_result in ocr_results
                for line in ocr_result[0]
            ]
            suitable_sizes = Converter._get_suitable_sizes(detected_sizes)

        for page_num, pdf_page in enumerate(pdf_document):
            pdf_page_width = pdf_page.rect.width
            pdf_page_height = pdf_page.rect.height

            # Create a new slide with blank layout
            slide_layout = pptx_output.slide_layouts[6]  # Blank layout
            slide = pptx_output.slides.add_slide(slide_layout)

            page_content = all_contents[page_num]
            drawings: list[dict] = page_content["drawings"]
            drawing_bboxes: list[pymupdf.Rect] = [
                drawing["rect"] for drawing in drawings
            ]
            images: list[dict] = page_content["images"]
            image_bboxes = [pymupdf.Rect(image["bbox"]) for image in images]
            xref_smask_map = page_content["xref_smask_map"]
            all_shapes = drawings + images
            all_bboxes = drawing_bboxes + image_bboxes

            # Sort indices of shapes by area in descending order
            indices = Converter._sort_shapes_by_area(all_bboxes)

            # Add the shapes to the slide. Smaller shapes will appear in front.
            for index in indices:
                shape = all_shapes[index]
                bbox = all_bboxes[index]
                # The shape is a drawing
                if index < len(drawings):
                    # Do not add svg characters to the slides
                    if scanned_document and Converter._drawing_is_character(shape):
                        continue
                    self._add_drawing_to_slide(
                        shape, bbox, pdf_page_width, pdf_page_height, slide
                    )
                # The shape is an image
                else:
                    xref = shape["xref"]
                    # This is an inline image
                    if xref == 0:
                        continue
                    # Do not add pure text images to slide
                    if scanned_document and self._image_is_pure_text(
                        pdf_document, xref
                    ):
                        continue
                    smask = xref_smask_map[xref]
                    self._add_image_to_slide(pdf_document, shape, smask, bbox, slide)

            # Iterate through text blocks and add them to the slide
            # NOTE: We add text blocks to slides AFTER we add images in order that they appear in front of the images.
            text_blocks: list[dict] = page_content["text_blocks"]
            for text_block in text_blocks:
                self._add_text_block_to_slide(text_block, slide)

            if scanned_document:
                page_ocr_results = ocr_results[page_num]
                # NOTE: The index 0 is necessary
                for line_num, ocr_line in enumerate(page_ocr_results[0]):
                    fontsize = suitable_sizes[
                        sum(len(page_ocr[0]) for page_ocr in ocr_results[:page_num])
                        + line_num
                    ]
                    self._add_ocr_to_slide(ocr_line, fontsize, slide)

    @staticmethod
    def _sort_shapes_by_area(all_bboxes: list[pymupdf.Rect]) -> list[int]:
        """Sort indices of shapes by area in descending order."""
        indices = sorted(
            range(len(all_bboxes)),
            key=lambda i: all_bboxes[i].get_area(),
            reverse=True,
        )
        return indices

    @staticmethod
    def _get_bbox(rect: list[list[float]]) -> list[float]:
        tl, tr, br, bl = rect
        left = min(tl[0], bl[0])
        top = min(tl[1], tr[1])
        right = max(tr[0], br[0])
        bottom = max(br[1], tr[1])
        return left, top, right, bottom

    @staticmethod
    def _get_ocr_detected_font_size(ocr_line: list) -> float:
        """Calculate the detected font size as the minimum of bbox height and width, clamped between 1 and 64."""

        # Since the input pixmap has dpi=300 rather than dpi=72 by default, we need to set the scaling factor as 72/300 = 0.24
        scaling_factor = 0.24

        rect = ocr_line[0]
        # Get bounding box coordinates
        left, top, right, bottom = Converter._get_bbox(rect)
        left *= scaling_factor
        top *= scaling_factor
        right *= scaling_factor
        bottom *= scaling_factor

        height = bottom - top
        width = right - left

        # Clamp the fontsize between 1 and 64 to avoid extreme fontsizes
        fontsize = np.clip(min(height, width), 1, 64)
        return fontsize

    @staticmethod
    def _gmm_bic_score(estimator, X):
        """Callable to pass to GridSearchCV that will use the BIC score.

        Reference: https://scikit-learn.org/stable/auto_examples/mixture/plot_gmm_selection.html#sphx-glr-auto-examples-mixture-plot-gmm-selection-py
        """
        # Make it negative since GridSearchCV expects a score to maximize
        return -estimator.bic(X)

    @staticmethod
    def _get_suitable_sizes(detected_sizes: list[float]) -> list[float]:
        """Determine the suitable font sizes using GMM clustering.

        Reference: https://scikit-learn.org/stable/auto_examples/mixture/plot_gmm_selection.html#sphx-glr-auto-examples-mixture-plot-gmm-selection-py
        """

        # Keep the detected sizes as they are, otherwise an error will be raised
        if len(detected_sizes) < 5:
            return detected_sizes

        # Convert the list into a numpy column vector
        X = np.asarray(detected_sizes).reshape((-1, 1))

        # Use grid search to determine the optimal number of components
        param_grid = {
            "n_components": range(1, 7),
        }
        grid_search = GridSearchCV(
            GaussianMixture(init_params="k-means++"),
            param_grid=param_grid,
            scoring=Converter._gmm_bic_score,
        )
        grid_search.fit(X)

        # Calculate the suitable sizes as the predicted cluster means
        best_estimator = grid_search.best_estimator_
        labels = best_estimator.predict(X)
        suitable_sizes = list(map(lambda i: best_estimator.means_[i], labels))
        return suitable_sizes

    @staticmethod
    def _transform_image(
        base_image: pymupdf.Pixmap,
        bbox: pymupdf.Rect,
        transformation_matrix: pymupdf.Matrix,
    ) -> bytes:
        """Transform the image according to the transformation matrix and return the transformed image in bytes."""

        # Convert the colorspace to RGB explicitly
        base_image = pymupdf.Pixmap(pymupdf.Colorspace(pymupdf.CS_RGB), base_image)

        # Add a fully opaque alpha channel if the channel is not present
        if not base_image.alpha:
            alpha = bytearray(255 for _ in range(base_image.w * base_image.h))
            base_image = pymupdf.Pixmap(base_image, 1)
            base_image.set_alpha(alpha)

        # Convert the base image to a Pillow Image object
        image_data = base_image.samples
        pil_image = Image.frombytes("RGBA", (base_image.w, base_image.h), image_data)

        # Define a scaling factor to enhance the output quality
        scaling_factor = 4

        # Calculate the matrix that transforms the original image to the displayed image
        shrink = pymupdf.Matrix(
            scaling_factor / base_image.w, 0, 0, scaling_factor / base_image.h, 0, 0
        )
        translation = pymupdf.Matrix(1, 0, 0, 1, -bbox[0], -bbox[1])
        transformation_matrix = shrink * transformation_matrix * translation
        transformation_matrix.invert()

        # Obtain the transformed image
        pil_image = pil_image.transform(
            (
                scaling_factor * int(bbox[2] - bbox[0]),
                scaling_factor * int(bbox[3] - bbox[1]),
            ),
            Image.AFFINE,
            # The parameters in PIL have a different order than in PyMuPDF
            (
                transformation_matrix.a,
                transformation_matrix.c,
                transformation_matrix.e,
                transformation_matrix.b,
                transformation_matrix.d,
                transformation_matrix.f,
            ),
            resample=Image.Resampling.BICUBIC,
        )

        # Convert Pillow image back to bytes
        output_bytes = io.BytesIO()
        pil_image.save(output_bytes, format="PNG")
        return output_bytes.getvalue()

    @staticmethod
    def _drawing_is_character(drawing: dict) -> bool:
        """Determine whether the drawing is a character using heuristics."""

        command_length_threshold = 6
        return all(
            [
                not drawing["closePath"],
                drawing["type"] == "f",
                len(drawing["items"]) > command_length_threshold,
            ]
        )

    def _image_is_pure_text(self, pdf_doc: pymupdf.Document, image_xref: dict) -> bool:
        """Use the layout engine to determine whether the image is pure text."""

        base_image = pymupdf.Pixmap(pdf_doc, image_xref)
        image_bytes = base_image.tobytes()
        layout_result = self.layout_engine(image_bytes)

        # The image is likely decorative. It does not contain text.
        if not layout_result:
            return False

        # Any area type other than figure is considered as text
        for area in layout_result:
            if area["type"] == "Figure":
                return False
        return True

    @staticmethod
    def _validate_arguments(input_file_path: str, output_file_path: str) -> None:
        if not os.path.isfile(input_file_path):
            raise FileNotFoundError(f"The file {input_file_path} does not exist.")

        input_file_extension = os.path.splitext(input_file_path)[1]
        if input_file_extension != ".pdf":
            raise ValueError(
                f"Invalid input file extension. Expected a .pdf file, but got {input_file_extension}"
            )

        output_file_extension = os.path.splitext(output_file_path)[1]
        if output_file_extension != ".pptx":
            raise ValueError(
                f"Invalid output file extension. Expected a .pptx file, but got {output_file_extension}"
            )
