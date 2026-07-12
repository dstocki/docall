from pathlib import Path
import shutil

import fitz  # PyMuPDF
import streamlit as st
from PIL import Image


WORK_DIR = Path("splitter_work")
PAGES_DIR = WORK_DIR / "pages"
OUTPUT_DIR = Path("split_output")


def clear_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


def render_pdf_pages(pdf_path: Path, output_dir: Path) -> list[Path]:
    clear_dir(output_dir)

    doc = fitz.open(pdf_path)
    image_paths = []

    for page_index, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2))
        img_path = output_dir / f"page_{page_index:04d}.png"
        pix.save(img_path)
        image_paths.append(img_path)

    return image_paths


def starts_to_ranges(starts: list[int], total_pages: int) -> list[tuple[int, int]]:
    starts = sorted(set(starts))

    if not starts or starts[0] != 1:
        starts = [1] + starts

    ranges = []

    for i, start in enumerate(starts):
        end = starts[i + 1] - 1 if i + 1 < len(starts) else total_pages
        ranges.append((start, end))

    return ranges


def split_pdf(pdf_path: Path, ranges: list[tuple[int, int]], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    source_doc = fitz.open(pdf_path)

    for idx, (start, end) in enumerate(ranges, start=1):
        new_doc = fitz.open()

        # PyMuPDF używa indeksów od 0
        new_doc.insert_pdf(
            source_doc,
            from_page=start - 1,
            to_page=end - 1,
        )

        output_path = output_dir / f"{pdf_path.stem}_invoice_{idx:03d}_pages_{start}-{end}.pdf"
        new_doc.save(output_path)
        new_doc.close()


def main():
    st.set_page_config(page_title="PDF Invoice Splitter", layout="wide")

    st.title("PDF Invoice Splitter")
    st.write("Wgraj jeden PDF z wieloma fakturami i zaznacz strony, na których zaczyna się nowa faktura.")

    uploaded_file = st.file_uploader("Wgraj PDF", type=["pdf"])

    if uploaded_file is None:
        return

    WORK_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    pdf_path = WORK_DIR / uploaded_file.name

    if "last_uploaded_file" not in st.session_state:
        st.session_state.last_uploaded_file = None

    if st.session_state.last_uploaded_file != uploaded_file.name:
        clear_dir(PAGES_DIR)
        pdf_path.write_bytes(uploaded_file.read())
        image_paths = render_pdf_pages(pdf_path, PAGES_DIR)

        st.session_state.last_uploaded_file = uploaded_file.name
        st.session_state.image_paths = [str(p) for p in image_paths]
    else:
        image_paths = [Path(p) for p in st.session_state.image_paths]
    total_pages = len(image_paths)

    st.success(f"Wczytano PDF: {uploaded_file.name}, liczba stron: {total_pages}")

    st.subheader("Zaznacz początki faktur")

    col_a, col_b = st.columns(2)

    with col_a:
        if st.button("Zaznacz wszystkie strony"):
            for i in range(1, total_pages + 1):
                st.session_state[f"start_page_{i}"] = True
            st.rerun()

    with col_b:
        if st.button("Zaznacz tylko pierwszą stronę"):
            for i in range(2, total_pages + 1):
                st.session_state[f"start_page_{i}"] = False
            st.session_state["start_page_1"] = True
            st.rerun()

    selected_starts = []

    cols_per_row = 3

    for row_start in range(0, total_pages, cols_per_row):
        cols = st.columns(cols_per_row)

        for offset, col in enumerate(cols):
            page_num = row_start + offset + 1

            if page_num > total_pages:
                continue

            img_path = image_paths[page_num - 1]

            with col:
                checked = st.checkbox(
                    f"Strona {page_num} = start faktury",
                    value=(page_num == 1),
                    key=f"start_page_{page_num}",
                )

                if checked:
                    selected_starts.append(page_num)

                image = Image.open(img_path)
                st.image(image, caption=f"Strona {page_num}", width='stretch')

    ranges = starts_to_ranges(selected_starts, total_pages)

    st.subheader("Zakresy faktur")

    for idx, (start, end) in enumerate(ranges, start=1):
        st.write(f"Faktura {idx}: strony {start}-{end}")

    if st.button("Podziel PDF"):
        clear_dir(OUTPUT_DIR)
        split_pdf(pdf_path, ranges, OUTPUT_DIR)

        st.success(f"Gotowe. Pliki zapisane w folderze: {OUTPUT_DIR.resolve()}")

        for file in sorted(OUTPUT_DIR.glob("*.pdf")):
            st.write(file.name)


if __name__ == "__main__":
    main()