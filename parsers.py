import io
import re
from collections import defaultdict
import pandas as pd
import pypdf
from pdfminer.layout import LAParams, LTTextBox
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator

MONTH_MAP = {
    'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'MAY': '05', 'JUN': '06',
    'JUL': '07', 'AUG': '08', 'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12'
}

def unlock_pdf_bytes(file_bytes, password=""):
    """Unlocks encrypted PDFs and permission locks in memory."""
    try:
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        if reader.is_encrypted:
            decrypted = False
            try:
                if reader.decrypt("") != 0:
                    decrypted = True
            except Exception:
                pass
            
            if not decrypted and password:
                try:
                    if reader.decrypt(password) != 0:
                        decrypted = True
                except Exception:
                    pass
            
            if not decrypted:
                return None, "⚠️ This PDF is password-protected. Please enter your PDF password and try again."
            
            writer = pypdf.PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            unenc_out = io.BytesIO()
            writer.write(unenc_out)
            return unenc_out.getvalue(), None
    except Exception:
        pass
    return file_bytes, None

def parse_td_chequing_pdf(file_bytes, year=2026):
    """Accurately extracts withdrawals and deposits, separating merged numbers and dates."""
    rsrcmgr = PDFResourceManager()
    laparams = LAParams(line_margin=0.1)
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    
    words = []
    fp = io.BytesIO(file_bytes)
    for page_idx, page in enumerate(PDFPage.get_pages(fp)):
        interpreter.process_page(page)
        layout = device.get_result()
        for element in layout:
            if isinstance(element, LTTextBox):
                for text_line in element:
                    txt = text_line.get_text().strip()
                    if txt:
                        words.append((page_idx, text_line.bbox[1], text_line.bbox[0], txt))
                        
    clustered = defaultdict(list)
    for p, y, x, txt in sorted(words, key=lambda item: (item[0], -item[1], item[2])):
        found = False
        for (cp, cy) in clustered:
            if cp == p and abs(cy - y) <= 4.0:
                clustered[(cp, cy)].append((x, txt))
                found = True
                break
        if not found:
            clustered[(p, y)].append((x, txt))
            
    txns = []
    date_regex = re.compile(r'(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s*(\d{2})', re.IGNORECASE)

    for (p, cy) in sorted(clustered.keys(), key=lambda k: (k[0], -k[1])):
        items = sorted(clustered[(p, cy)], key=lambda x: x[0])
        row_txt = " ".join([t for _, t in items])
        if any(h in row_txt for h in ['START ING BALANCE', 'CLOSING BALANCE', 'Account /Transact', 'Overdraft', 'Branch No.', 'UNLIMITED', 'Descript ion']):
            continue
            
        date_item = None
        desc_parts = []
        withdrawal = None
        deposit = None
        
        for x, t in items:
            clean_t = t.replace(" ", "").replace(",", "").replace("$", "")
            
            m_date = date_regex.search(clean_t)
            if m_date:
                date_item = (m_date.group(1).upper(), m_date.group(2))
                clean_t = clean_t[:m_date.start()] + clean_t[m_date.end():]
                
            m_amt = re.search(r'^-?(\d+\.\d{2})', clean_t)
            if m_amt:
                amt_val = float(m_amt.group(1))
                if 250 <= x < 340:
                    withdrawal = amt_val
                elif 340 <= x < 420:
                    deposit = amt_val
            elif x < 250:
                desc_parts.append(t)
                
        if date_item and (withdrawal is not None or deposit is not None):
            m_str, d_str = date_item
            iso_date = f"{year}-{MONTH_MAP[m_str]}-{d_str}"
            desc = " ".join(desc_parts).replace("_", " ").strip()
            
            desc_tokens = desc.split()
            cleaned_tokens = []
            for tok in desc_tokens:
                if not cleaned_tokens or tok != cleaned_tokens[-1]:
                    cleaned_tokens.append(tok)
            desc = " ".join(cleaned_tokens)
            
            amt = withdrawal if withdrawal is not None else deposit
            is_dep = deposit is not None
            txns.append({
                "Date": iso_date,
                "Description": desc,
                "Amount": amt,
                "is_deposit": is_dep
            })
    return pd.DataFrame(txns)

def parse_td_visa_pdf(file_bytes, year=2026):
    """Extracts TD Visa credit card transactions."""
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    date_re = re.compile(
        r'^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{1,2})\s+(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{1,2})\s+(-?\$?[\d,]+\.\d{2})(.*)$',
        re.IGNORECASE
    )
    txns = []
    for page in reader.pages:
        txt = page.extract_text() or ""
        for line in txt.split('\n'):
            m = date_re.match(line.strip())
            if m:
                m1, d1, _, _, amt_str, desc = m.groups()
                amt = float(amt_str.replace("$", "").replace(",", ""))
                iso_date = f"{year}-{MONTH_MAP[m1.upper()]}-{int(d1):02d}"
                txns.append({
                    "Date": iso_date,
                    "Description": desc.strip(),
                    "Amount": amt,
                    "is_deposit": amt < 0
                })
    return pd.DataFrame(txns)
