from argostranslate.tags import translate_tags
from argostranslate.translate import ITranslation
from docx import Document

from argostranslatefiles.formats.abstract_xml import AbstractXml


class Docx(AbstractXml):
    supported_file_extensions = ['.docx']

    def translate(self, underlying_translation: ITranslation, file_path: str):
        outzip_path = self.get_output_path(underlying_translation, file_path)
        document = Document(file_path)
        for paragraph in document.paragraphs:
            if len(paragraph.runs) == 1:
                paragraph.runs[0].text = underlying_translation.translate(paragraph.runs[0].text)
            elif len(paragraph.runs) == 0:
                continue
            else:
                first_run_style = paragraph.runs[0].style
                paragraph.text = underlying_translation.translate(paragraph.text)
                paragraph.runs[0].style = first_run_style

        document.save(outzip_path)

        return outzip_path
