"""
Module 1: TXT Parsing Engine
- Page-break detection
- Extract Title Page, Appearances, Indices, Exhibits, Ending, Disclosure, Certificate
"""

import re
from dateutil import parser as dateparser


COURT_HEADING_STOP_PATTERN = re.compile(
    r"(PLAINTIFF|DEFENDANT|VS\.?.|CIVIL ACTION|CASE NO\.?|FILE NO\.?|APPELLANT|RESPONDENT)",
    re.IGNORECASE,
)


def _looks_like_court_heading_line(line: str) -> bool:
    """Return True when a line resembles an uppercase court heading."""

    if not line:
        return False
    has_court = "COURT" in line.upper()
    alpha = sum(ch.isalpha() for ch in line)
    lower = sum(ch.islower() for ch in line)
    upperish = alpha and (lower / alpha) <= 0.25
    return has_court and upperish


def extract_court_heading_from_lines(lines):
    """Capture a court heading block (one or more uppercase lines) from sequential lines."""

    for idx, line in enumerate(lines):
        if not line:
            continue
        if _looks_like_court_heading_line(line.strip()):
            heading = [line.strip()]
            for follow in lines[idx + 1 :]:
                stripped = follow.strip()
                if not stripped:
                    break
                if COURT_HEADING_STOP_PATTERN.search(stripped):
                    break
                alpha = sum(ch.isalpha() for ch in stripped)
                lower = sum(ch.islower() for ch in stripped)
                if alpha and (lower / alpha) > 0.25:
                    break
                heading.append(stripped)
            return " ".join(heading)
    return ""

class TXTData:
    def __init__(self):
        self.pages = {}               # page_number -> list of lines
        self.raw = ""                 # full raw text
        self.title = {}               # dict of extracted title page fields
        self.appearances = {}         # dict
        self.indices = {}             # dict
        self.exhibits = {}            # dict
        self.ending = {}              # dict
        self.disclosure = {}          # dict
        self.certificate = {}         # dict


class TXTParser:
    PAGE_NUMBER_REGEX = re.compile(r'^\s*(\d{1,4})\s*$')

    def load(self, path):
        data = TXTData()
        with open(path, "r", errors="ignore", encoding="utf-8") as f:
            data.raw = f.read()
        lines = data.raw.splitlines()
        self._split_pages(lines, data)
        self._parse_title_page(data)
        self._parse_appearances(data)
        self._parse_indices(data)
        self._parse_exhibits(data)
        self._parse_ending(data)
        self._parse_disclosure(data)
        self._parse_certificate(data)
        return data

    def _split_pages(self, lines, data):
        """Split TXT into pages using isolated page numbers."""
        pages = {}
        current_page = 1
        pages[current_page] = []

        for line in lines:
            if self.PAGE_NUMBER_REGEX.match(line.strip()):
                num = int(line.strip())
                if num not in pages:
                    current_page = num
                    pages[current_page] = []
                continue
            pages[current_page].append(line)

        data.pages = pages

    def _parse_title_page(self, data):
        p1 = data.pages.get(1, [])
        joined = "\n".join(p1)

        def find(pattern):
            m = re.search(pattern, joined, re.IGNORECASE)
            return m.group(0).strip() if m else ""

        # Basic patterns
        data.title["court_heading"] = extract_court_heading_from_lines(p1)
        data.title["case_number"] = find(r"(202\d.*|20\d{2}.*|CIVIL ACTION.*|FILE NO.*)")
        data.title["case_style"] = self._find_case_style(joined)

        # Witness name detection
        data.title["witness_name"] = self._extract_above_witness_title(p1)

        # Job title detection
        data.title["job_title"], data.title["job_adjectives"] = self._extract_job_title(p1)

        # Date extraction
        data.title["date"] = self._extract_date(joined)

        # Time
        m = re.search(r"(\d{1,2}:\d{2}\s*(AM|PM|a\.m\.|p\.m\.))", joined)
        data.title["start_time"] = m.group(1) if m else ""

        # Location - bracketed parenthetical
        m = re.search(r"\(.*?\)", joined)
        data.title["location"] = m.group(0) if m else ""

        # Reporter
        m = re.search(r"Reported by\s+(.+)", joined, re.IGNORECASE)
        data.title["resource"] = m.group(1).strip() if m else ""

    def _find_case_style(self, txt):
        # Grab block around Plaintiff/Defendant
        m = re.search(r"(.+Plaintiff.+vs\.?.+Defendant)", txt, re.IGNORECASE)
        return m.group(1).strip() if m else ""

    def _extract_above_witness_title(self, lines):
        # Find job title lines then look above/below
        for i, line in enumerate(lines):
            if re.search(r"Deposition of|Hearing|Arbitration|Examination|Trial", line, re.IGNORECASE):
                # look up 5 lines
                window = lines[max(0,i-5):i+5]
                for w in window:
                    if re.search(r"[A-Z][A-Za-z'\-]+\s+[A-Z][A-Za-z'\-]+", w):
                        return w.strip().strip(",")
        return ""

    def _extract_job_title(self, lines):
        # Option C: title = line above witness
        for i, line in enumerate(lines):
            if re.search(r"[A-Z][A-Za-z'\-]+\s+[A-Z][A-Za-z'\-]+", line) and "Plaintiff" not in line:
                # line above likely job title
                if i > 0:
                    jt = lines[i-1].strip()
                    adjectives = self._extract_adjectives(jt)
                    return jt, adjectives
        return "", []

    def _extract_adjectives(self, line):
        adjs = []
        possible = ["Remote","Hybrid","30(b)(6)","Videoconference","Videotaped","Excerpt","Continuation","Confidential"]
        for a in possible:
            if a.lower() in line.lower():
                adjs.append(a)
        return adjs

    def _extract_date(self, txt):
        m = re.search(r"(January|February|March|April|May|June|July|August|September|October|November|December)[^,]+,\s*\d{4}", txt)
        if m:
            try:
                dt = dateparser.parse(m.group(0))
                return dt.strftime("%B %d, %Y")
            except:
                return m.group(0)
        return ""

    def _parse_appearances(self, data):
        # Page 2 assumed
        p2 = "\n".join(data.pages.get(2, []))
        data.appearances["heading_present"] = bool(re.search(r"APPEARANCES", p2, re.IGNORECASE))

        # Extract "On behalf of"
        roles = re.findall(r"On behalf of the ([A-Za-z]+)", p2, re.IGNORECASE)
        data.appearances["sides"] = roles

        # Attorneys & firms
        attys = re.findall(r"([A-Z][A-Za-z'.\-]+,\s*Esq\.?)", p2)
        data.appearances["attorneys"] = attys

    def _parse_indices(self, data):
        p3 = data.pages.get(3, [])
        joined = "\n".join(p3)

        # Index to Examinations
        m = re.search(r"INDEX TO EXAMINATIONS(.+?)(INDEX TO EXHIBITS|$)", joined, re.IGNORECASE|re.DOTALL)
        data.indices["exam_index_block"] = m.group(1) if m else ""

        # Extract exam page numbers
        ex_pages = re.findall(r"Examination.+?(\d+)", joined, re.IGNORECASE)
        data.indices["exam_pages"] = [int(x) for x in ex_pages]

        # Index to Exhibits
        m = re.search(r"INDEX TO EXHIBITS(.+)$", joined, re.IGNORECASE|re.DOTALL)
        data.indices["exhibit_index_block"] = m.group(1) if m else ""

        ex_pages = re.findall(r"Exhibit\s+(\d+).+?(\d+)", joined, re.IGNORECASE)
        data.indices["exhibit_pages"] = [(int(e), int(p)) for e,p in ex_pages]

    def _parse_exhibits(self, data):
        # Scan all pages for parentheticals like "(Exhibit 1 ...)"
        exhibits_found = {}
        for pg, lines in data.pages.items():
            for line in lines:
                m = re.search(r"\(.*Exhibit\s+(\d+).*?\)", line, re.IGNORECASE)
                if m:
                    num = int(m.group(1))
                    exhibits_found.setdefault(num, []).append(pg)
        data.exhibits["locations"] = exhibits_found

    def _parse_ending(self, data):
        # Scan last 5 pages for end time, signature
        last_pgs = sorted(data.pages.keys())[-5:]
        block = "\n".join(sum((data.pages[p] for p in last_pgs), []))

        m = re.search(r"(concluded|adjourned|dismissed|suspended).*?(\d{1,2}:\d{2}\s*(AM|PM|a\.m\.|p\.m\.))", block, re.IGNORECASE)
        data.ending["end_time"] = m.group(2) if m else ""

        m = re.search(r"(signature.*?reserved|signature.*?waived)", block, re.IGNORECASE)
        data.ending["signature"] = m.group(0) if m else ""

    def _parse_disclosure(self, data):
        alltxt = data.raw
        m = re.search(r"(DISCLOSURE.*?)(CERTIFICATE|STATE OF)", alltxt, re.IGNORECASE|re.DOTALL)
        block = m.group(1) if m else ""
        data.disclosure["block"] = block
        data.disclosure["date"] = self._extract_date(block)

        m = re.search(r"CCR|Court Reporter", block, re.IGNORECASE)
        data.disclosure["resource"] = m.group(0) if m else ""

    def _parse_certificate(self, data):
        alltxt = data.raw
        m = re.search(r"(CERTIFICATE.*)", alltxt, re.IGNORECASE|re.DOTALL)
        block = m.group(1) if m else ""
        data.certificate["block"] = block

        m = re.search(r"(STATE OF|COUNTY OF).+", block)
        data.certificate["court"] = m.group(0) if m else ""

        data.certificate["date"] = self._extract_date(block)

        m = re.search(r"(CCR.+|Court Reporter.+)", block)
        data.certificate["resource"] = m.group(0) if m else ""


# ============================
# Module 2: PDF Parsing Engine
# ============================

import re
from PyPDF2 import PdfReader

class PDFParser:
    def load(self, path):
        data = {}
        text = ""
        first_page_text = ""
        try:
            reader = PdfReader(path)
            for idx, page in enumerate(reader.pages):
                page_text = page.extract_text() or ""
                if idx == 0:
                    first_page_text = page_text
                text += page_text
        except:
            return data

        primary_text = first_page_text or text
        raw_lines = [ln.rstrip() for ln in primary_text.splitlines()]
        # Normalize
        normalized = re.sub(r'\s+', ' ', primary_text)

        data["court_heading"] = self._extract_court_heading(primary_text, raw_lines)
        data["case_number"]   = self._find(normalized, r"(CIVIL ACTION FILE NO\.?|FILE NO\.?)\s*[#:]*\s*([A-Za-z0-9\-\/\.]+)", group=2)
        data["case_style"]    = self._find(normalized, r".+?,\s*Plaintiff.*?v\.?.+?,\s*Defendant", flags=re.IGNORECASE)
        data["witness_name"]  = self._find(normalized, r"Deposition of\s+(.+?)(?=[,\.])", group=1)
        data["date"]          = self._find(normalized, r"(January|February|March|April|May|June|July|August|September|October|November|December)[^,]*,\s*\d{4}")
        data["start_time"]    = self._find(normalized, r"(\d{1,2}:\d{2}\s*(AM|PM|a\.m\.|p\.m\.))", group=1)
        data["location"]      = self._find(normalized, r"Location:\s*(.*?)\s{2,}", group=1)

        return data

    def _find(self, text, pattern, group=0, flags=0):
        m = re.search(pattern, text, flags)
        if not m:
            return ""
        return m.group(group).strip()

    def _extract_court_heading(self, page_text, raw_lines):
        """Run court heading detection once to avoid duplicated heuristics."""

        heading = extract_court_heading_from_lines(raw_lines)
        if heading:
            return heading
        return self._find_heading_block(page_text)

    def _find_heading_block(self, text):
        """Fallback to grab a heading block from the first page text while stopping at party metadata."""
        snippet = text[:1200]
        # If no explicit heading line found, try a regex anchored near COURT and stop at party labels
        m = re.search(r"(IN\s+THE|THE)\s+[^\n]*COURT[^\n]*", snippet, re.IGNORECASE)
        if not m:
            return ""
        start = m.start()
        tail_lines = [ln.strip() for ln in snippet[start:].splitlines() if ln.strip()]
        collected = []
        for ln in tail_lines:
            if COURT_HEADING_STOP_PATTERN.search(ln):
                break
            alpha = sum(ch.isalpha() for ch in ln)
            lower = sum(ch.islower() for ch in ln)
            if alpha and (lower / alpha) > 0.25 and collected:
                break
            collected.append(ln)
        return " ".join(collected)



# ============================
# Module 3: RB Loader & Mapping
# ============================

import pandas as pd
import re
from urllib.parse import urlparse, parse_qs

class RBJobData:
    def __init__(self, jobno):
        self.job_number = jobno
        self.fields = {}

    def set(self, key, value):
        self.fields[key] = value or ""

class RBLoader:
    def __init__(self):
        # Hardcoded URLs
        self.rb_pull_url = "https://docs.google.com/spreadsheets/d/1s04mN8nh-n7rFZG8yPJTTKU1IGmtjZ6MBh_ggpiD3dU/export?format=csv&gid=445432281"
        self.firms_url   = "https://docs.google.com/spreadsheets/d/1s04mN8nh-n7rFZG8yPJTTKU1IGmtjZ6MBh_ggpiD3dU/export?format=csv&gid=1546586553"
        self.exhibits_url= "https://docs.google.com/spreadsheets/d/1k5qnuRlRa04-PAuCderyikq9-MFsFZzoTa9hnDwE_gk/export?format=csv&gid=0"

    def load_sheet(self, url):
        try:
            df = pd.read_csv(url, dtype=str).fillna("")
            return df
        except:
            return pd.DataFrame()

    def load_all(self):
        df_rb = self.load_sheet(self.rb_pull_url)
        df_firms = self.load_sheet(self.firms_url)
        df_ex = self.load_sheet(self.exhibits_url)
        return df_rb, df_firms, df_ex

    def get_job_data(self, jobno):
        df_rb, df_firms, df_ex = self.load_all()
        job = RBJobData(jobno)

        # Find job row
        if "JobNo" in df_rb.columns:
            row = df_rb[df_rb["JobNo"].astype(str) == str(jobno)]
            if not row.empty:
                r = row.iloc[0]
                for col in df_rb.columns:
                    job.set(col, r.get(col,""))

        # Firm mapping
        firm_name = job.fields.get("OrderingFirm","").strip().lower()
        if firm_name and "Firm name" in df_firms.columns:
            fm = df_firms[df_firms["Firm name"].str.strip().str.lower() == firm_name]
            if not fm.empty:
                f = fm.iloc[0]
                for col in df_firms.columns:
                    job.set("Firm_"+col.replace(" ","_"), f.get(col,""))

        # Exhibits mapping
        if "JobNo" in df_ex.columns:
            ex = df_ex[df_ex["JobNo"].astype(str) == str(jobno)]
            if not ex.empty:
                e = ex.iloc[0]
                for col in df_ex.columns:
                    job.set("Ex_"+col.replace(" ","_"), e.get(col,""))

        return job


# ============================
# Module 4A: Comparison Utilities (Clean Version)
# ============================

import re
from datetime import datetime
from dateutil import parser as dateparser

class ComparisonResult:
    def __init__(self, index, group, field, txt, pdf, rb, status, notes=""):
        self.index = index
        self.group = group
        self.field = field
        self.txt = txt or ""
        self.pdf = pdf or ""
        self.rb = rb or ""
        self.status = status
        self.notes = notes

def normalize_ws(s):
    if not s:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def normalize_case(s):
    return normalize_ws(s).upper()

def similar(a, b):
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    a = a.lower()
    b = b.lower()
    def bigrams(x):
        return {x[i:i+2] for i in range(len(x)-1)}
    A = bigrams(a)
    B = bigrams(b)
    if not A and not B:
        return 1.0
    if not A or not B:
        return 0.0
    return len(A & B) / len(A | B)

def parse_time_val(s):
    if not s:
        return None
    try:
        return dateparser.parse(s).time()
    except:
        return None

def time_diff_minutes(t1, t2):
    if not t1 or not t2:
        return None
    dt1 = datetime.combine(datetime.today(), t1)
    dt2 = datetime.combine(datetime.today(), t2)
    return abs((dt1 - dt2).total_seconds()) / 60.0

def compare_text(index, group, field, txt, pdf, rb, notes=""):
    t_txt = normalize_ws(txt)
    t_pdf = normalize_ws(pdf)
    t_rb  = normalize_ws(rb)

    if not t_txt:
        status = "MISSING" if (t_pdf or t_rb) else "MISSING"
        return ComparisonResult(index, group, field, t_txt, t_pdf, t_rb, status, notes)

    EXACT_THR = 0.90
    PARTIAL_THR = 0.60

    results = []

    if t_pdf:
        s = similar(t_txt, t_pdf)
        if s >= EXACT_THR:
            results.append("PDF_EXACT")
        elif s >= PARTIAL_THR:
            results.append("PDF_PARTIAL")
        else:
            results.append("PDF_NONE")
    else:
        results.append("PDF_MISSING")

    if t_rb:
        s = similar(t_txt, t_rb)
        if s >= EXACT_THR:
            results.append("RB_EXACT")
        elif s >= PARTIAL_THR:
            results.append("RB_PARTIAL")
        else:
            results.append("RB_NONE")
    else:
        results.append("RB_MISSING")

    if all(r.endswith("EXACT") for r in results if not r.endswith("MISSING")):
        status = "EXACT_MATCH"
    elif any(r.endswith("EXACT") or r.endswith("PARTIAL") for r in results):
        if any(r.endswith("NONE") for r in results):
            status = "NO_MATCH"
        else:
            status = "PARTIAL_MATCH"
    else:
        status = "NO_MATCH"

    return ComparisonResult(index, group, field, t_txt, t_pdf, t_rb, status, notes)

def compare_time(index, group, field, txt, pdf, rb, notes=""):
    T_txt = parse_time_val(txt)
    T_pdf = parse_time_val(pdf)
    T_rb  = parse_time_val(rb)

    if not T_txt:
        status = "MISSING"
        return ComparisonResult(index, group, field, txt, pdf, rb, status, notes)

    results = []

    if T_pdf:
        d = time_diff_minutes(T_txt, T_pdf)
        results.append("PDF_OK" if d is not None and d <= 5 else "PDF_NONE")
    else:
        results.append("PDF_MISSING")

    if T_rb:
        d = time_diff_minutes(T_txt, T_rb)
        results.append("RB_EXACT" if d is not None and d <= 0.01 else "RB_NONE")
    else:
        results.append("RB_MISSING")

    if all(r in ("PDF_OK", "RB_EXACT") for r in results if not r.endswith("MISSING")):
        status = "EXACT_MATCH"
    elif any(r in ("PDF_OK", "RB_EXACT") for r in results):
        status = "PARTIAL_MATCH" if not any(r.endswith("NONE") for r in results) else "NO_MATCH"
    else:
        status = "NO_MATCH"

    return ComparisonResult(index, group, field, txt, pdf, rb, status, notes)

def strict_page_match(actual_page, expected_page):
    return actual_page == expected_page

# ============================
# Module 4B: Run All Comparisons
# ============================

def run_all_comparisons(txt_data, pdf_data, rb_data):
    results = []

    def c(index, group, field, txt, pdf, rb):
        return compare_text(index, group, field, txt, pdf, rb)

    def ct(index, group, field, txt, pdf, rb):
        return compare_time(index, group, field, txt, pdf, rb)

    # 1-9 Title Page
    results.append(c(1,"Title Page","Court Heading", txt_data.title.get("court_heading"), pdf_data.get("court_heading"), rb_data.fields.get("Case Court/County","")))
    results.append(c(2,"Title Page","Case Number", txt_data.title.get("case_number"), pdf_data.get("case_number"), rb_data.fields.get("CaseNo","")))
    results.append(c(3,"Title Page","Case Style", txt_data.title.get("case_style"), pdf_data.get("case_style"), rb_data.fields.get("CaseFullName","")))
    results.append(c(4,"Title Page","Job Title", txt_data.title.get("job_title"), "", rb_data.fields.get("TaskType","")))
    results.append(c(5,"Title Page","Witness Name", txt_data.title.get("witness_name"), pdf_data.get("witness_name"), rb_data.fields.get("Witness","")))
    results.append(c(6,"Title Page","Date", txt_data.title.get("date"), pdf_data.get("date"), rb_data.fields.get("JobDate","")))
    results.append(ct(7,"Title Page","Start Time", txt_data.title.get("start_time"), pdf_data.get("start_time"), rb_data.fields.get("ActualStartTime","")))
    results.append(c(8,"Title Page","Location", txt_data.title.get("location"), pdf_data.get("location"), rb_data.fields.get("JobLocAddress","")))
    results.append(c(9,"Title Page","Resource", txt_data.title.get("resource"), "", rb_data.fields.get("Resource","")))

    # 10-12 Appearances
    heading = "Yes" if txt_data.appearances.get("heading_present") else "No"
    s = "EXACT_MATCH" if txt_data.appearances.get("heading_present") else "NO_MATCH"
    results.append(ComparisonResult(10,"Appearances","Appearances Heading Present", heading,"","",s))

    sides = txt_data.appearances.get("sides", [])
    cap = normalize_case(txt_data.title.get("case_style",""))
    ok = any(normalize_case(x) in cap for x in sides)
    s = "EXACT_MATCH" if ok else "NO_MATCH"
    results.append(ComparisonResult(11,"Appearances","On Behalf of Side Matches Title Page", ", ".join(sides),"","",s))

    atty = txt_data.appearances.get("attorneys", [])
    s = "EXACT_MATCH" if atty else "NO_MATCH"
    results.append(ComparisonResult(12,"Appearances","Attorney Names / Contact Info Present", ", ".join(atty),"","",s))

    # 13-18 Indices
    results.append(c(13,"Indices","Witness Name (Index Page Verification)", txt_data.title.get("witness_name"),"", rb_data.fields.get("Witness","")))

    blk = txt_data.indices.get("exam_index_block","")
    s = "EXACT_MATCH" if blk else "NO_MATCH"
    results.append(ComparisonResult(14,"Indices","Index to Examinations Present", "Yes" if blk else "No","","",s))

    ex_pages = txt_data.indices.get("exam_pages", [])
    s = "EXACT_MATCH" if ex_pages else "NO_MATCH"
    results.append(ComparisonResult(15,"Indices","Index to Examinations Page Numbers Correct", str(ex_pages),"","",s))

    blk = txt_data.indices.get("exhibit_index_block","")
    s = "EXACT_MATCH" if blk else "NO_MATCH"
    results.append(ComparisonResult(16,"Indices","Index to Exhibits Present", "Yes" if blk else "No","","",s))

    ex_pages = txt_data.indices.get("exhibit_pages", [])
    s = "EXACT_MATCH" if ex_pages else "NO_MATCH"
    results.append(ComparisonResult(17,"Indices","Index to Exhibits Page Numbers Correct", str(ex_pages),"","",s))

    exloc = txt_data.exhibits.get("locations", {})
    s = "EXACT_MATCH" if exloc else "NO_MATCH"
    results.append(ComparisonResult(18,"Indices","Exhibit Parentheticals Present", str(exloc),"","",s))

    # 19-22 Ending
    results.append(c(19,"Ending","Job Title Heading (Ending Section)", txt_data.title.get("job_title"),"", rb_data.fields.get("TaskType","")))
    results.append(c(20,"Ending","Date (Ending Section)", txt_data.title.get("date"),"", rb_data.fields.get("JobDate","")))
    results.append(ct(21,"Ending","End Time", txt_data.ending.get("end_time"),"", rb_data.fields.get("ActualEndTime","")))

    sig = txt_data.ending.get("signature")
    s = "EXACT_MATCH" if sig else "NO_MATCH"
    results.append(ComparisonResult(22,"Ending","Signature Parenthetical", sig or "","","",s))

    # 23-25 Disclosure
    blk = txt_data.disclosure.get("block","")
    s = "EXACT_MATCH" if blk else "NO_MATCH"
    results.append(ComparisonResult(23,"Disclosure","Disclosure Page Present", "Yes" if blk else "No","","",s))

    results.append(c(24,"Disclosure","Disclosure Date", txt_data.disclosure.get("date"),"", rb_data.fields.get("JobDate","")))
    results.append(c(25,"Disclosure","Disclosure Resource", txt_data.disclosure.get("resource"),"", rb_data.fields.get("Resource","")))

    # 26-29 Certificate
    blk = txt_data.certificate.get("block","")
    s = "EXACT_MATCH" if blk else "NO_MATCH"
    results.append(ComparisonResult(26,"Certificate","Certificate Heading Present", "Yes" if blk else "No","","",s))

    crt = txt_data.certificate.get("court","")
    s = "EXACT_MATCH" if crt else "NO_MATCH"
    results.append(ComparisonResult(27,"Certificate","Certificate Court Subheading Present", crt or "","","",s))

    results.append(c(28,"Certificate","Certificate Date", txt_data.certificate.get("date"),"", rb_data.fields.get("JobDate","")))

    results.append(c(29,"Certificate","Certificate Resource", txt_data.certificate.get("resource"),"", rb_data.fields.get("Resource","")))

    return results

# ============================
# Module 4C: Result Aggregation
# ============================

def organize_results(results):
    exact = []
    partial = []
    no = []
    missing = []

    for r in results:
        if r.status == "EXACT_MATCH":
            exact.append(r)
        elif r.status == "PARTIAL_MATCH":
            partial.append(r)
        elif r.status == "NO_MATCH":
            no.append(r)
        else:
            missing.append(r)

    keyfn = lambda x: x.index
    exact.sort(key=keyfn)
    partial.sort(key=keyfn)
    no.sort(key=keyfn)
    missing.sort(key=keyfn)

    return {
        "EXACT": exact,
        "PARTIAL": partial,
        "NO": no,
        "MISSING": missing
    }

# ============================
# Module 5A: PDF Report Setup
# ============================

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import LETTER
from reportlab.lib import colors

def _shorten(s, limit=120):
    if not s:
        return ""
    s = str(s).strip()
    return s if len(s) <= limit else s[:limit] + "..."

class PDFReportBuilder:
    def __init__(self, path):
        self.path = path
        self.styles = getSampleStyleSheet()
        self.title_style = self.styles["Title"]
        self.header_style = self.styles["Heading2"]
        self.normal = self.styles["Normal"]
        self.value_style = ParagraphStyle(
            "Values",
            parent=self.normal,
            leftIndent=18,
            textColor=colors.red
        )

    def build(self, summary_groups, all_results):
        story = []
        self._add_summary_page(story, summary_groups)
        story.append(PageBreak())
        self._add_detail_pages(story, all_results)
        doc = SimpleDocTemplate(self.path, pagesize=LETTER,
                                rightMargin=40, leftMargin=40,
                                topMargin=40, bottomMargin=40)
        doc.build(story)
        return self.path

    def _add_summary_page(self, story, groups):
        pass  # Implemented in Module 5B

    def _add_detail_pages(self, story, all_results):
        pass  # Implemented in Module 5C

# ============================
# Module 5B: Summary Page Builder
# ============================

def _section_header(builder, story, title):
    story.append(Paragraph(title, builder.header_style))
    story.append(Spacer(1, 6))

def _summary_line(builder, story, result):
    txt = f"{result.index}. {result.group} – {result.field}"
    story.append(Paragraph(txt, builder.normal))
    story.append(Spacer(1, 3))

def PDFReportBuilder__add_summary_page(self, story, groups):
    story.append(Paragraph("QC Summary", self.title_style))
    story.append(Spacer(1, 12))

    order_titles = [
        ("EXACT", "Exact Matches"),
        ("PARTIAL", "Partial Matches"),
        ("NO", "No Matches"),
        ("MISSING", "Missing Items")
    ]

    for key, label in order_titles:
        _section_header(self, story, label)
        section = groups.get(key, [])
        if not section:
            story.append(Paragraph("None.", self.normal))
            story.append(Spacer(1, 6))
            continue
        for r in section:
            _summary_line(self, story, r)
        story.append(Spacer(1, 10))

PDFReportBuilder._add_summary_page = PDFReportBuilder__add_summary_page

# ============================
# Module 5C: Detail Pages Builder
# ============================

def PDFReportBuilder__add_detail_pages(self, story, all_results):
    for r in sorted(all_results, key=lambda x: x.index):
        header = f"{r.index}. {r.group} – {r.field}: {r.status}"
        story.append(Paragraph(header, self.header_style))

        val_txt = _shorten(r.txt)
        val_pdf = _shorten(r.pdf)
        val_rb  = _shorten(r.rb)

        parts = []
        if val_txt: parts.append(f"TXT: {val_txt}")
        if val_pdf: parts.append(f"PDF: {val_pdf}")
        if val_rb:  parts.append(f"RB: {val_rb}")

        body = " | ".join(parts)
        if r.notes:
            body += f"  Notes: {r.notes}"

        story.append(Paragraph(body, self.value_style))
        story.append(Spacer(1, 12))

PDFReportBuilder._add_detail_pages = PDFReportBuilder__add_detail_pages

# ============================
# Module 5D: Report Orchestrator
# ============================

import os


class QCSummary:
    """Lightweight container returned by :func:`run_qc`."""

    def __init__(self, job_number, witness_name, counts):
        self.job_number = job_number or ""
        self.witness_name = witness_name or ""
        self.counts = counts or {}

def generate_qc_pdf_report(output_path, results_grouped, all_results):
    builder = PDFReportBuilder(output_path)
    return builder.build(results_grouped, all_results)


# ============================
# Module 6: Orchestrator (run_qc)
# ============================

import glob


def _find_first(path_pattern):
    matches = glob.glob(path_pattern)
    return matches[0] if matches else ""


def _derive_job_number(job_folder, txt_path):
    # Prefer folder name digits, fall back to TXT filename digits.
    for candidate in (job_folder, os.path.basename(txt_path)):
        m = re.search(r"(\d+)", os.path.basename(candidate))
        if m:
            return m.group(1)
    return ""


def run_qc(job_folder):
    """Run the full QC pipeline for the given job folder.

    Returns a tuple of (QCSummary, report_path).
    """

    if not os.path.isdir(job_folder):
        raise FileNotFoundError(f"Job folder not found: {job_folder}")

    txt_path = _find_first(os.path.join(job_folder, "*.txt"))
    if not txt_path:
        raise FileNotFoundError("No TXT transcript found in job folder")

    pdf_path = _find_first(os.path.join(job_folder, "*.pdf"))

    txt_parser = TXTParser()
    txt_data = txt_parser.load(txt_path)

    pdf_parser = PDFParser()
    pdf_data = pdf_parser.load(pdf_path) if pdf_path else {}

    job_number = _derive_job_number(job_folder, txt_path)
    rb_loader = RBLoader()
    rb_data = rb_loader.get_job_data(job_number) if job_number else RBJobData("")

    all_results = run_all_comparisons(txt_data, pdf_data, rb_data)
    grouped = organize_results(all_results)

    counts = {
        "EXACT_MATCH": sum(1 for r in all_results if r.status == "EXACT_MATCH"),
        "PARTIAL_MATCH": sum(1 for r in all_results if r.status == "PARTIAL_MATCH"),
        "NO_MATCH": sum(1 for r in all_results if r.status == "NO_MATCH"),
        "MISSING": sum(1 for r in all_results if r.status not in {"EXACT_MATCH", "PARTIAL_MATCH", "NO_MATCH"}),
    }

    report_path = os.path.join(job_folder, "QC_Report.pdf")
    generate_qc_pdf_report(report_path, grouped, all_results)

    summary = QCSummary(job_number, txt_data.title.get("witness_name"), counts)
    return summary, report_path
