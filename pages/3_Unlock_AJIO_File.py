import io
import re
import zipfile

import streamlit as st

st.set_page_config(page_title="Unlock AJIO File", page_icon="🔓", layout="centered")

st.title("🔓 Unlock AJIO File")
st.write(
    "AJIO / marketplace Excel file upload karo. Sheet protection aur workbook "
    "structure lock hat jayega, aur data, formulas, formatting, dropdowns, "
    "comments sab jaisa hai waisa hi rahega."
)

SHEET_PROT = re.compile(r"<sheetProtection\b[^>]*?/>|<sheetProtection\b[^>]*?>.*?</sheetProtection>", re.S)
BOOK_PROT = re.compile(r"<workbookProtection\b[^>]*?/>|<workbookProtection\b[^>]*?>.*?</workbookProtection>", re.S)
PROTECTED_RANGES = re.compile(r"<protectedRanges\b.*?</protectedRanges>", re.S)


def unlock_workbook(file_bytes: bytes):
    """Return (unlocked_bytes, report). Only protection tags are removed."""
    zin = zipfile.ZipFile(io.BytesIO(file_bytes))
    names = zin.namelist()

    # Password-to-open (encrypted) files are not zip files / have EncryptedPackage
    if "EncryptedPackage" in names or "xl/workbook.xml" not in names:
        raise ValueError(
            "Ye file open karne ke liye password-protected (encrypted) hai ya valid .xlsx nahi hai. "
            "Pehle Excel me password dalkar open karo aur bina password ke save karo."
        )

    # map sheet file -> sheet name
    wb_xml = zin.read("xl/workbook.xml").decode("utf8", "ignore")
    sheet_names = re.findall(r'<sheet\b[^>]*?name="([^"]*)"', wb_xml)

    report = {"sheets_unlocked": [], "workbook_structure_unlocked": False}
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        sheet_idx = 0
        for item in zin.infolist():
            data = zin.read(item.filename)
            fn = item.filename

            if re.match(r"xl/(worksheets|chartsheets)/[^/]+\.xml$", fn):
                text = data.decode("utf8")
                new_text, n = SHEET_PROT.subn("", text)
                new_text = PROTECTED_RANGES.sub("", new_text)
                if n:
                    label = fn.split("/")[-1]
                    report["sheets_unlocked"].append(label)
                    data = new_text.encode("utf8")

            elif fn == "xl/workbook.xml":
                text = data.decode("utf8")
                new_text, n = BOOK_PROT.subn("", text)
                if n:
                    report["workbook_structure_unlocked"] = True
                    data = new_text.encode("utf8")

            zout.writestr(item, data)

    report["sheet_names"] = sheet_names
    return out_buf.getvalue(), report


uploaded = st.file_uploader(
    "Excel file upload karo", type=["xlsx", "xlsm"], accept_multiple_files=True
)

if uploaded:
    for up in uploaded:
        st.divider()
        st.subheader(up.name)
        try:
            new_bytes, rep = unlock_workbook(up.getvalue())
        except Exception as e:  # noqa: BLE001
            st.error(str(e))
            continue

        if not rep["sheets_unlocked"] and not rep["workbook_structure_unlocked"]:
            st.info("Is file me koi protection nahi mili. File pehle se unlocked hai.")
        else:
            st.success(
                f"Unlock ho gaya: {len(rep['sheets_unlocked'])} sheet(s) ki protection hata di"
                + (", workbook structure lock bhi hata diya." if rep["workbook_structure_unlocked"] else ".")
            )

        base, ext = up.name.rsplit(".", 1)
        mime = (
            "application/vnd.ms-excel.sheet.macroEnabled.12"
            if ext.lower() == "xlsm"
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.download_button(
            "⬇️ Unlocked file download karo",
            data=new_bytes,
            file_name=f"{base}_unlocked.{ext}",
            mime=mime,
            key=f"dl_{up.name}",
        )
else:
    st.caption("Ek ya zyada files ek saath upload kar sakte ho.")
