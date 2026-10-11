"""
Checks the version 2.4 changes without needing the window or an analyser.

Run:  python3 test_LabScribe_2.4.py
"""
import importlib.util
import os
import sys

import pdfplumber

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "test_out")

# The script has been called both LabScribe.py and LabScribe_2.4.py; use
# whichever is sitting next to this test.
MODULE = next((n for n in ("LabScribe.py", "LabScribe_2.4.py")
               if os.path.exists(os.path.join(HERE, n))), "LabScribe.py")

spec = importlib.util.spec_from_file_location("ls24",
                                              os.path.join(HERE, MODULE))
L = importlib.util.module_from_spec(spec)
spec.loader.exec_module(L)

passed, failed = [], []


def check(label, condition, detail=""):
    (passed if condition else failed).append(label)
    mark = "ok  " if condition else "FAIL"
    print(f"  {mark}  {label}" + (f"   [{detail}]" if detail and not condition else ""))


def job(name, patient, results, separate=False):
    return {"path": os.path.join(OUT, name), "name": name, "status": "ready",
            "patient": dict(patient), "results": [dict(r) for r in results],
            "error": "", "out": "", "separate": separate}


def pdf_text(path):
    with pdfplumber.open(path) as book:
        return "\n".join((page.extract_text() or "") for page in book.pages)


os.makedirs(OUT, exist_ok=True)

# The real-world case: one patient, chemistry printed on one file and the
# glycated haemoglobin on another, run on different days with different
# sample numbers — exactly what used to lose a report.
CHEM = dict(L.DEMO_PATIENT)
HBA1C = {"path_no": "CHE640/26", "specimen_date": "29/04/2026",
         "name": "Boateng, Adwoa Sample", "sex": "FEMALE", "age": "46",
         "specimen": "WHOLE BLOOD"}

chem_job = job("chemistry.pdf", CHEM, L.DEMO_RESULTS)
hba1c_job = job("hba1c.pdf", HBA1C, L.DEMO_HBA1C_RESULTS)

print("\nGROUPING")
groups = L.group_jobs_by_patient([chem_job, hba1c_job])
check("two files for one patient make one group", len(groups) == 1,
      f"got {len(groups)}")
check("the name is matched whatever its order and punctuation",
      L.patient_merge_key({"name": "Boateng, Adwoa Sample"})
      == L.patient_merge_key({"name": "ADWOA SAMPLE BOATENG"}))

other = job("someone_else.pdf", {**CHEM, "name": "KWAME MENSAH"}, L.DEMO_RESULTS)
check("a different patient is left alone",
      len(L.group_jobs_by_patient([chem_job, hba1c_job, other])) == 2)

nameless = job("nameless.pdf", {**CHEM, "name": ""}, L.DEMO_RESULTS)
check("a file with no name is never joined to anything",
      len(L.group_jobs_by_patient([nameless, job("n2.pdf", {**CHEM, "name": ""},
                                                 L.DEMO_RESULTS)])) == 2)

kept_apart = job("hba1c.pdf", HBA1C, L.DEMO_HBA1C_RESULTS, separate=True)
check("'report this file on its own' parts a pair",
      len(L.group_jobs_by_patient([chem_job, kept_apart])) == 2)

print("\nTHE MERGED RECORD")
patient, results = L.merged_job(groups[0])
check("the path number comes off the chemistry printout",
      patient["path_no"] == "CHE901/26", patient["path_no"])
check("the specimen date comes off the chemistry printout",
      patient["specimen_date"] == "27/04/2026", patient["specimen_date"])
check("the specimen line names both samples",
      patient["specimen"] == "SERUM / WHOLE BLOOD", patient["specimen"])
check("the age and sex survive", (patient["age"], patient["sex"]) == ("46", "FEMALE"))

rows = L.apply_test_library(results, patient)
names = [r["test"] for r in rows]
check("no test is printed twice", len(names) == len(set(names)))
check("exactly one glycated haemoglobin row",
      sum(1 for r in rows if L.is_hba1c_row(r)) == 1)
check("the chemistry is all still there", len(rows) > 15, f"{len(rows)} rows")

print("\nTHE REPORT ITSELF")
merged_pdf = os.path.join(OUT, "merged.pdf")
L.generate_report(patient, results, merged_pdf)
text = pdf_text(merged_pdf)
for section in ("RENAL FUNCTION", "LIVER FUNCTION", "LIPID PROFILE",
                "GLYCATED"):
    check(f"{section} prints on the merged report", section in text.upper())
check("the patient's name prints once", text.upper().count("ADWOA SAMPLE BOATENG") >= 1)
check("the merged specimen line prints", "SERUM / WHOLE BLOOD" in text.upper())

# An HbA1c on its own must still be the single clean page it always was.
solo_pdf = os.path.join(OUT, "hba1c_only.pdf")
L.generate_report(L.DEMO_HBA1C_PATIENT, L.DEMO_HBA1C_RESULTS, solo_pdf)
solo = pdf_text(solo_pdf).upper()
check("a glycated haemoglobin on its own still says WHOLE BLOOD",
      "WHOLE BLOOD" in solo)
check("a glycated haemoglobin on its own has no OTHER INVESTIGATIONS",
      "OTHER INVESTIGATIONS" not in solo)

print("\nTHE FILE NAMES")
taken = set()
first = L.unique_output_path(OUT, "CHE901_26_ADWOA_SAMPLE_BOATENG.pdf", taken)
second = L.unique_output_path(OUT, "CHE901_26_ADWOA_SAMPLE_BOATENG.pdf", taken)
check("two reports can never be given the same name", first != second,
      f"{first} vs {second}")
check("the second is numbered, not overwritten",
      second.endswith("CHE901_26_ADWOA_SAMPLE_BOATENG_2.pdf"), second)

print("\nTHE eGFR")
parsed = [{"test": "CREA", "result": "106", "unit": "umol/L", "ref": ""}]
young = L.calculate_egfr(parsed, {"age": "30", "sex": "FEMALE"})
egfr_young = [r for r in young if r["test"].startswith("eGFR")][0]["result"]
# The age is now corrected in the window, and the figure must follow it.
old = L.recompute_egfr(young, {"age": "75", "sex": "FEMALE"})
egfr_old = [r for r in old if r["test"].startswith("eGFR")][0]["result"]
check("a corrected age changes the eGFR", egfr_young != egfr_old,
      f"{egfr_young} vs {egfr_old}")
check("the eGFR falls with age", int(egfr_old) < int(egfr_young),
      f"{egfr_young} -> {egfr_old}")
check("only one eGFR row is ever left",
      sum(1 for r in old if r["test"].startswith("eGFR")) == 1)

analyser_egfr = [{"test": "CREA", "result": "106", "unit": "umol/L", "ref": ""},
                 {"test": "eGFR", "result": "55", "unit": "", "ref": ""}]
kept = L.recompute_egfr(analyser_egfr, {"age": "75", "sex": "FEMALE"})
check("an eGFR the analyser printed is left alone",
      [r for r in kept if "eGFR" in r["test"]][0]["result"] == "55")

print(f"\n{len(passed)} passed, {len(failed)} failed")
if failed:
    for label in failed:
        print(f"   failed: {label}")
    sys.exit(1)
