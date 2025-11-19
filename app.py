import subprocess
from flask import Flask, request, send_file
import os
import uuid
import pythoncom
import win32com.client
import threading
import time
import traceback

# ===============================
# Kill any zombie Word processes
# ===============================
try:
    subprocess.call(["taskkill", "/IM", "WINWORD.EXE", "/F"], stderr=subprocess.DEVNULL)
except:
    pass

# ----------------------
# Configuration
# ----------------------
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
CLEANUP_INTERVAL = 60
FILE_MAX_AGE = 300

app = Flask(__name__)

# ----------------------
# Helpers
# ----------------------
def _paragraph_index_for_position(doc, pos):
    for idx, para in enumerate(doc.Paragraphs, 1):
        try:
            if para.Range.Start <= pos <= para.Range.End:
                return idx
        except:
            continue
    return 1 if doc.Paragraphs.Count >= 1 else None

def _safe_find_replace_range(rng, find_text, replace_text=""):
    try:
        find = rng.Find
        wdReplaceAll = 2
        find.Execute(FindText=find_text, ReplaceWith=replace_text, Replace=wdReplaceAll)
    except:
        pass

def _remove_invisible_chars_in_range(rng):
    invisible_chars = ["\u200b", "\ufeff"]
    for ch in invisible_chars:
        _safe_find_replace_range(rng, ch, "")

def _delete_empty_paragraphs_at_start(doc, max_iter=20):
    count = 0
    while doc.Paragraphs.Count >= 1 and count < max_iter:
        first_para = doc.Paragraphs(1)
        txt = first_para.Range.Text or ""
        if txt.strip() == "":
            try:
                first_para.Range.Delete()
            except:
                break
            count += 1
        else:
            break

# ----------------------
# UPDATE TOC FUNCTION
# ----------------------
def update_toc_word(input_path, output_path):
    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(
            input_path,
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
            Revert=False,
            OpenAndRepair=True
        )

        # SAFE ACCESS CHECK
        try:
            _ = doc.Content.Text
        except Exception as err:
            raise Exception(f"Document content unavailable: {err}")

        # REMOVE EXISTING TOCS
        while doc.TablesOfContents.Count > 0:
            try:
                doc.TablesOfContents(1).Delete()
            except:
                break

        # FIND TOC HEADING (look for "Table of Contents", "Contents", "TOC")
        toc_heading_para = None
        search_texts = ["table of contents", "toc", "contents"]

        for para in doc.Paragraphs:
            try:
                text = para.Range.Text.strip().lower()
                if any(k in text for k in search_texts):
                    toc_heading_para = para
                    break
            except:
                continue

        # CLEANUP REGION (before TOC heading)
        try:
            if toc_heading_para:
                clean_range = doc.Range(0, toc_heading_para.Range.Start)
            else:
                clean_range = doc.Range(0, min(2000, doc.Content.End))
        except:
            clean_range = doc.Range(0, min(2000, doc.Content.End))

        # REMOVE MANUAL PAGE BREAKS (^m), SECTION BREAKS (^b), invisible chars
        _safe_find_replace_range(clean_range, "^m", "")
        _safe_find_replace_range(clean_range, "^b", "")
        _remove_invisible_chars_in_range(clean_range)

        # DELETE EMPTY PARAGRAPHS AT START
        _delete_empty_paragraphs_at_start(doc, max_iter=50)

        # -------------------------
        # ENSURE TOC HEADING IS ON ITS OWN PAGE
        # -------------------------
        if toc_heading_para:
            heading_range = toc_heading_para.Range.Duplicate
            heading_range.Collapse(1)  # Collapse to start of heading
            if toc_heading_para.Range.Start > 0:
                heading_range.InsertBreak(7)  # 7 = wdPageBreak

        # -------------------------
        # DETERMINE TOC INSERT POSITION
        # -------------------------
        if toc_heading_para:
            # Insert TOC immediately after heading paragraph
            toc_insert_range = toc_heading_para.Range.Duplicate
            toc_insert_range.Collapse(0)  # Collapse to end
        else:
            toc_insert_range = doc.Range(0, 0)
            toc_insert_range.Collapse(0)

        # ADD TOC
        try:
            toc = doc.TablesOfContents.Add(
                toc_insert_range,
                UseHeadingStyles=True,
                UpperHeadingLevel=1,
                LowerHeadingLevel=3,
                RightAlignPageNumbers=True,
                IncludePageNumbers=True
            )
        except Exception:
            toc = doc.TablesOfContents.Add(toc_insert_range)

        # UPDATE TOC
        for t in doc.TablesOfContents:
            try:
                t.Update()
            except:
                continue

        # FINAL CLEANUP
        _delete_empty_paragraphs_at_start(doc, max_iter=10)

        # SAVE
        doc.SaveAs(output_path)

    finally:
        try:
            if doc:
                doc.Close(False)
        except:
            pass
        try:
            if word:
                word.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

# ----------------------
# FLASK ROUTE
# ----------------------
@app.route("/update-toc", methods=["POST"])
def update_toc():
    if "file" not in request.files:
        return {"error": "No file uploaded"}, 400

    file = request.files["file"]
    if file.filename == "":
        return {"error": "No selected file"}, 400

    input_path = os.path.join(DOWNLOAD_DIR, f"input_{uuid.uuid4().hex}.docx")
    output_path = os.path.join(DOWNLOAD_DIR, f"updated_{uuid.uuid4().hex}.docx")

    file.save(input_path)
    file.stream.close()
    time.sleep(0.1)

    try:
        update_toc_word(input_path, output_path)
    except Exception as e:
        return {"error": str(e), "traceback": traceback.format_exc()}, 500

    return send_file(
        output_path,
        as_attachment=True,
        download_name=f"updated_{file.filename}",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        conditional=False
    )

# ----------------------
# CLEANUP THREAD
# ----------------------
def cleanup_old_files():
    while True:
        now = time.time()
        for fname in os.listdir(DOWNLOAD_DIR):
            fpath = os.path.join(DOWNLOAD_DIR, fname)
            try:
                if os.path.isfile(fpath) and now - os.path.getmtime(fpath) > FILE_MAX_AGE:
                    os.remove(fpath)
            except:
                pass
        time.sleep(CLEANUP_INTERVAL)

threading.Thread(target=cleanup_old_files, daemon=True).start()

# ----------------------
# RUN SERVER
# ----------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True, threaded=False)
