""".doc to .docx conversion using Word COM.

Corresponds to the implicit "open any Word file" capability in VBA.
"""
from pathlib import Path


def ensure_docx(doc_path: Path, *, output_dir: Path | None = None) -> Path:
    """Ensure the given path is a .docx file.

    If the input is already .docx, return it unchanged.
    If the input is .doc, convert it via Word COM and return the .docx path.
    Raises FileNotFoundError if the input does not exist.
    Raises RuntimeError if COM conversion fails.
    """
    if not doc_path.exists():
        raise FileNotFoundError(f"Document not found: {doc_path}")

    if doc_path.suffix.lower() == ".docx":
        return doc_path

    if doc_path.suffix.lower() == ".doc":
        out_dir = output_dir or doc_path.parent
        out_path = out_dir / (doc_path.stem + ".docx")

        import win32com.client

        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(doc_path))
            # wdFormatXMLDocument = 16
            doc.SaveAs2(str(out_path), FileFormat=16)
            doc.Close(SaveChanges=False)
            doc = None
            word.Quit()
            word = None
            return out_path
        except Exception as e:
            try:
                if doc:
                    doc.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if word:
                    word.Quit()
            except Exception:
                pass
            raise RuntimeError(f"Failed to convert .doc to .docx: {e}") from e

    return doc_path
