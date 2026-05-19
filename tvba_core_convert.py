from pathlib import Path
import tempfile


def _make_conversion_path(doc_path: Path, output_dir: Path, *, unique: bool) -> Path:
    """Return a unique .docx path for converting a legacy .doc file."""
    if unique:
        tmp = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".docx",
            prefix=f"{doc_path.stem}_transvba_",
            dir=output_dir,
        )
        tmp_path = Path(tmp.name)
        tmp.close()
        tmp_path.unlink(missing_ok=True)
        return tmp_path

    return output_dir / (doc_path.stem + ".docx")


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
        use_unique_temp = output_dir is None
        out_dir = output_dir or doc_path.parent
        out_path = _make_conversion_path(doc_path, out_dir, unique=use_unique_temp)

        import win32com.client

        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(doc_path), ReadOnly=True, AddToRecentFiles=False)
            # wdFormatXMLDocument = 16
            doc.SaveAs2(str(out_path), FileFormat=16)
            return out_path
        except Exception as e:
            raise RuntimeError(f"Failed to convert .doc to .docx: {e}") from e
        finally:
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

    return doc_path
