# EV Allocation Report – Filename rules and column mapping

Use this as a reference for which filename triggers which format and how columns are mapped to the output file.

---

## Filename rules (order matters: first match wins)

| If filename **contains**   | Format key  | Used for mapping below |
| -------------------------- | ----------- | ---------------------- |
| Erickson                   | erickson           | ✓                      |
| Viva Smiles (anywhere in filename, case-insensitive) | hoang_viva_smiles  | ✓       |
| Ismiles (anywhere in filename, case-insensitive)     | hoang_ismiles      | ✓       |
| Kates                      | kates              | ✓                      |
| Montefiore                 | montefiore  | ✓                      |
| ortho                      | ortho       | ✓                      |
| SL Evening or SL_Evening   | sl_evening  | ✓                      |
| SL medicaid or SL_medicaid | sl_medicaid | ✓                      |

Matching is **case-insensitive**.

---

## Column mapping by format

### erickson

_(filename contains "Erickson")_

| Output column       | Input column / value    |
| ------------------- | ----------------------- |
| System              | **"Edge"** (fixed value for all rows) |
| Office/Doctor Name  | **"Dr. Erickson"** (fixed value for all rows) |
| Source              | **"Evening"** (fixed value for all rows) |
| Reference           | **"MCD"** (fixed value for all rows) |
| Received Date       | **Today's date (MM/DD/YYYY)**        |
| Location/EntityCode | Patient Office          |
| Appointment         | Appointment Next Date   |
| Patients Name       | Patient Full Name       |
| DOB                 | Patient Birthdate       |
| Patient ID/Chart#   | Patient Primary Code    |
| Insurance           | Insurance Company Name  |
| Policy ID           | InsDetail Subscriber Id |
| Carrier Phone       | Insurance Company Phone |
| Subscriber Name     | InsDetal Subscriber     |
| Subscriber DOB      | Subscriber BirthDate    |

---

### hoang_viva_smiles

_(filename contains **"Viva Smiles"** anywhere, case-insensitive — e.g. Raw Dr. Hoang (Viva Smiles) 02.04.2026 (7to9))_

| Output column     | Input column / value                  |
| ----------------- | ------------------------------------- |
| System            | **"Dolphin"** (fixed value for all rows) |
| Office/Doctor Name | **"Dr. Hoang Viva Smiles"** (fixed value for all rows) |
| Source            | **"Evening"** (fixed value for all rows) |
| Reference         | **"MCD"** when Insurance Company Billing Center Name = "Humana Healthy"; **"Commercial"** when = "Aetna HMO" or "Aetna DHMO"; otherwise blank |
| Received Date     | **Today's date (MM/DD/YYYY)**         |
| Appointment       | Next Appointment Date                 |
| Patients Name     | Patient's Name (Last First)           |
| DOB               | Patient's BirthDate                   |
| Patient ID/Chart# | Patient's ID                          |
| Insurance         | Insurance Company Billing Center Name |
| Policy ID         | Subscriber ID                         |

---

### hoang_ismiles

_(filename contains **"Ismiles"** anywhere, case-insensitive — e.g. Raw Dr. Hoang Ismiles (Main) 02.04.2026 (7to9).csv)_

Same column mapping as Hoang Viva Smiles (see above).

| Output column     | Input column                          |
| ----------------- | ------------------------------------- |
| Appointment       | Next Appointment Date                 |
| Patients Name     | Patient's Name (Last First)           |
| DOB               | Patient's BirthDate                   |
| Patient ID/Chart# | Patient's ID                          |
| Insurance         | Insurance Company Billing Center Name |
| Policy ID         | Subscriber ID                         |

_Office/Doctor Name: blank (only filled for SL Evening files)._

---

### kates

_(filename contains "Kates")_

| Output column       | Input column    |
| ------------------- | --------------- |
| Location/EntityCode | Tx Location     |
| Appointment         | Scheduled Appts |
| Patients Name       | Full Name       |
| DOB                 | DOB             |
| Patient ID/Chart#   | Patient ID      |
| Insurance           | Payor           |
| Policy ID           | Member ID       |

---

### montefiore

_(filename contains "Montefiore")_

**Reference** is derived from **Insurance Company Billing Center Name**: **"MCD"** when the billing center name is in the configured MCD list (`EV_ALLOCATION_MONTEFIORE_MCD_BILLING_CENTERS`); **"Commercial"** when in the Commercial list (`EV_ALLOCATION_MONTEFIORE_COMMERCIAL_BILLING_CENTERS`) or when = "Aetna HMO" or "Aetna DHMO"; otherwise blank.

| Output column       | Input column                              |
| ------------------- | ----------------------------------------- |
| Location/EntityCode | Next Appointment Including Today Location |
| Reference           | **"MCD"** if Insurance Company Billing Center Name is in MCD list; **"Commercial"** if in Commercial list or Aetna HMO/DHMO |
| Appointment         | Next Appointment Including Today Date     |
| Patients Name       | Patient's Name (Last First)               |
| DOB                 | Patient's BirthDate                       |
| Patient ID/Chart#   | Patient ID                                |
| Insurance           | Insurance Company Billing Center Name     |
| Policy ID           | Subscriber ID                             |
| Subscriber Name     | Subscriber Name                           |
| Subscriber DOB      | Subscriber Birthdate                      |

---

### ortho

_(filename contains "ortho")_

| Output column       | Input column / logic |
| ------------------- | -------------------- |
| **Office/Doctor Name** | **"Dr. Mansman"** if Entity Code = "FREDORMD"; **"Dr. Susan Park"** if Entity Code = "SYRACUSE" or "NTHSYRNY"; otherwise **blank**. |
| Location/EntityCode | Entity Code     |
| Appointment         | Next Appt       |
| Patients Name       | Patient         |
| DOB                 | Pats Birth Date |
| Patient ID/Chart#   | Chart           |
| Insurance           | Carrier         |
| Policy ID           | Insured ID      |
| Carrier Phone       | Phone           |

---

### sl_evening

_(filename contains "SL Evening" or "SL_Evening")_

**Office/Doctor Name is filled only for this file type** (from Office Name). All other file types leave Office/Doctor Name blank.

| Output column       | Input column / logic                      |
| ------------------- | ----------------------------------------- |
| Office/Doctor Name  | Office Name                               |
| Location/EntityCode | _(blank for sl_evening)_                  |
| Reference           | **"MCD"** when Insurance Company Billing Center Name is in the MCD list (`EV_ALLOCATION_MONTEFIORE_MCD_BILLING_CENTERS`), or when Carrier Name = "United Healthcare"; **"Commercial"** when Carrier Name is in the Commercial list (`EV_ALLOCATION_MONTEFIORE_COMMERCIAL_BILLING_CENTERS`); otherwise **"MCD"** |
| Appointment         | Future Appt                               |
| Patients Name       | Pats Last Name + Pats First Name (merged) |
| DOB                 | Pats Birth Date                           |
| Insurance           | Carrier Name                              |
| Policy ID           | Pol Employee SSN ID                       |
| Carrier Phone       | Carrier Phone                             |
| Subscriber Name     | Emp Name                                  |
| Subscriber DOB      | Employee Birth Date                       |

---

### sl_medicaid

_(filename contains "SL medicaid" or "SL_medicaid")_

| Output column           | Input column / logic                                                                                                                                                                                                                                        |
| ----------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Office/Doctor Name**  | **Same value as Location/EntityCode** (see below).                                                                                                                                                                                                           |
| **Location/EntityCode** | **Special:** Find rows where **Pats First Name** contains `Office Name: <office_name>`, extract **office_name** from that text and put it in Location/EntityCode; use that value for that row and all following rows until the next "Office Name: ..." row. |
| Appointment             | Future Appt                                                                                                                                                                                                                                                 |
| Patients Name           | Pats Last Name + Pats First Name (merged); on “Office Name:” header rows, only Pats Last Name is used.                                                                                                                                                      |
| DOB                     | Pats Birth Date                                                                                                                                                                                                                                             |
| Insurance               | Carrier Name                                                                                                                                                                                                                                                |
| Policy ID               | Pol Employee SSN ID                                                                                                                                                                                                                                         |
| Carrier Phone           | Carrier Phone                                                                                                                                                                                                                                               |
| Subscriber Name         | Emp name last, First Need to merge                                                                                                                                                                                                                          |
| Subscriber DOB          | Employee Birth Date                                                                                                                                                                                                                                         |

---

## Department and Practice ID (all formats)

**Department** and **Practice ID** are filled from a lookup keyed by **(Office/Doctor Name, Reference)**. Reference is the value set per format (e.g. "MCD" or "Commercial"). Example: when Office/Doctor Name = "FREDPEDO" and Reference = "MCD", Department = "Medicaid 180" and Practice ID = "6002". The lookup table is `EV_ALLOCATION_DEPARTMENT_PRACTICE_LOOKUP` in code; matching is case-insensitive.

---

## Output file columns (fixed order)

1. System
2. Office/Doctor Name
3. Practice ID
4. Location/EntityCode
5. Department
6. Source
7. Received Date
8. Appointment
9. Reference
10. Patients Name
11. DOB
12. Patient ID/Chart#
13. Group/Employer
14. Insurance
15. Policy ID
16. Carrier Phone
17. Status
18. Comments
19. Pre Auth Status
20. Subscriber Name
21. Subscriber DOB
22. Zip Code
23. Rep
24. Agent
25. Remark
26. Work Date
27. QC Agent
28. QC Comments
29. QC Date Work

Any output column not listed in a format’s mapping is left blank for that file type.
