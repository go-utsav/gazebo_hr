# UK Holiday Records Compliance 2026

**Source:** Employment Rights Act 2025 · Effective **6 April 2026**  
**Organisation:** Gazebo Fine Food – Employee Reporting  
**Prepared:** April 2026

---

## 1. Legal context — why this matters

From **6 April 2026**, all UK employers **must** keep adequate records of employee holiday entitlement and holiday pay.

| Requirement | Detail |
|-------------|--------|
| **Retention** | Minimum **6 years** — annual leave taken, pay calculations, and carry-over |
| **Enforcement** | Failure is a **criminal offence** with **unlimited fines** (no longer civil-only) |
| **Inspector** | **Fair Work Agency (FWA)** launches **7 April 2026** — can inspect premises and seize documents |
| **Retrospective scope** | FWA has retrospective powers; records should cover leave from **December 2025 onward** |

---

## 2. Objective

Ensure Gazebo Fine Food is fully compliant with the Employment Rights Act 2025 by **accurately recording, calculating, and retaining** all employee holiday entitlement and pay data — **before 6 April 2026**.

---

## 3. What we must do (task)

1. **Extract** attendance data from the **ClockRite C400** clocking machine (via ClockRite Desktop export).
2. **Record** each employee's **days worked per week** (all employees are **regular-hours** — no irregular/zero-hours at Gazebo).
3. **Apply** the regular-hours statutory formula: `days per week × 5.6`.
4. **Produce** a clean, auditable record per employee.
5. **Store** records securely for **6 years**.

---

## 4. Submission vs retention

**There is no proactive submission to government.**

- Records are **not** sent to HMRC, FWA, or any body by default.
- Records must **exist** and be **producible on request** within **28 days**.

### Who may request records

- Fair Work Agency (from April 2026)
- Employment tribunals
- Employees (Subject Access Request)
- HMRC (if linked to National Minimum Wage audit)

### Acceptable format

- Any readable, auditable format (Excel / CSV recommended)
- **One clear record per employee**
- Cloud backup strongly advised
- Treat security like payroll — **do not delete when staff leave**

---

## 5. Holiday calculation formula

**Scope:** All Gazebo employees are **regular-hours workers**. The irregular/zero-hours (12.07%) method does not apply.

### Regular-hours workers

```
Days worked per week × 5.6 = Annual leave days
```

| Example | Calculation | Result |
|---------|-------------|--------|
| Full-time (5 days/week) | 5 × 5.6 | **28 days** |
| Part-time (3 days/week) | 3 × 5.6 | **16.8 days** |

**Rules:**

- Statutory minimum = **5.6 weeks**
- Most employers give 28 days (20 + 8 bank holidays)
- **Cannot round down** — round up to nearest **0.5 day**

---

## 6. End-to-end workflow

| Step | Action |
|------|--------|
| **1. Export** | ClockRite Desktop → **Reports → Export** as CSV or Excel. Fields: Name, Date, Clock-In, Clock-Out, Hours |
| **2. Days per week** | Record each employee's contracted days per week (all regular-hours) |
| **3. Calculate** | Apply `days per week × 5.6` (§5) |
| **4. Leave taken** | Cross-reference approved leave (HR system or paper forms); subtract from entitlement |
| **5. Holiday pay** | Verify pay rates include **regular overtime and commission**, not just basic pay |
| **6. Store** | Save final record with employee name, dates, entitlement, taken, pay rate. Back up securely; label by tax year |

**Export frequency:** weekly or monthly (repeat process monthly / at year-end).

---

## 7. Data source — ClockRite C400

| Attribute | Detail |
|-----------|--------|
| Device | Facial recognition attendance terminal |
| Software | ClockRite Desktop on Windows PC |
| Built-in | Holiday Status report |
| Export path | ClockRite Software → Reports → Export |
| Export format | CSV or Excel (`.xlsx`) |
| Export fields | Name, Date, Clock-In, Clock-Out, Hours |

**Automation target:** A Python script reads the ClockRite export and applies the correct holiday formula to generate the compliance report.

---

## 8. Required output — per employee, per year

### 8.1 Core fields (example)

| Employee | Days/Week | Entitlement | Leave Taken | Remaining | Pay Rate Used |
|----------|-----------|-------------|-------------|-----------|---------------|
| Sarah Ahmed | 5 | 28 days | 15 days | 13 days | £14.50/hr |
| Priya Nair | 3 | 16.8 days | 10 days | 6.8 days | £16.00/hr |

### 8.2 Additional mandatory fields

Each record must also include:

- Leave year **start** and **end** dates
- Any **carry-over** and the reason
- **Payment for unused leave** on termination
- **Evidence of encouragement to use leave** (e.g. emails)

---

## 9. Functional requirements (system)

### 9.1 Employee data

- [ ] All employees treated as **regular-hours** (no zero-hours / irregular classification)
- [ ] Store **days worked per week** per employee (e.g. 5 full-time, 3 part-time)

### 9.2 Entitlement calculation

- [ ] `days_per_week × 5.6`, round up to nearest 0.5 day
- [ ] Support leave-year boundaries (start/end dates)

### 9.3 Leave tracking

- [ ] Record leave taken (from HR / manual input)
- [ ] Compute remaining balance: `entitlement − taken`
- [ ] Track carry-over with reason

### 9.4 Holiday pay

- [ ] Store pay rate used for holiday calculations
- [ ] Pay rate must reflect regular overtime and commission where applicable

### 9.5 Import / export

- [ ] Parse ClockRite CSV/Excel exports (clock-in/out, hours)
- [ ] Generate per-employee compliance report (Excel/CSV)
- [ ] Label exports by tax year

### 9.6 Retention & audit

- [ ] Store records for minimum **6 years**
- [ ] Do not delete on employee departure
- [ ] Secure backup (cloud)
- [ ] Produce records on request within **28 days**

---

## 10. Action checklist (manual / pre-automation)

1. [ ] Find ClockRite software on office PC and open it
2. [ ] Reports → Export → save sample CSV/Excel
3. [ ] List all employees with days worked per week
4. [ ] Apply `days per week × 5.6` per person
5. [ ] Record leave actually taken (HR records or paper forms)
6. [ ] Verify holiday pay rates include regular overtime/commission
7. [ ] Save completed records to secure, backed-up location
8. [ ] Set calendar reminder to repeat monthly / at year-end

---

## 11. Relation to existing `gazebo_hr` app

The weekly payroll app already parses **ClockRite Paid Hours (Inc Absence) Summary** exports and extracts `AnnualHoliday` hours from Excel columns H/L (`weekly/payroll_service.py`). That covers **hours already paid as annual leave** in payroll, not the full statutory compliance record described here.

**Gaps to build:**

| Area | Current state | Compliance need |
|------|---------------|-----------------|
| Days per week | Not in models | Per-employee contracted days/week (all regular-hours) |
| Statutory entitlement | Not calculated | `days_per_week × 5.6` |
| Leave balance | Partial (paid hours only) | Entitlement earned vs taken vs remaining |
| Carry-over / termination pay | Not tracked | Required fields |
| 6-year retention | Not implemented | Persistent store + backup |
| Encouragement evidence | Not tracked | Document storage / reference |
| Raw clock export | Weekly summary only | ClockRite export for attendance / leave reconciliation |

---

## 12. Success criteria (“done”)

- [ ] Every employee has one auditable holiday record per leave year
- [ ] Entitlement calculated with `days_per_week × 5.6`
- [ ] Leave taken reconciled against HR approvals
- [ ] Holiday pay rate documented (incl. OT/commission where relevant)
- [ ] Carry-over, termination payment, and leave-year dates recorded
- [ ] Records stored securely for 6+ years
- [ ] Process repeatable monthly / year-end from ClockRite export
- [ ] Report producible within 28 days if inspected by FWA or employee SAR

---

## 13. References

- Employment Rights Act 2025 (effective 6 April 2026)
- Employment Rights Act 2024 reform (12.07% irregular-hours method — not used at Gazebo)
- ClockRite Support: **01246 267715**
- Internal PDF: `.cursor/plans/UK_Holiday_Compliance_2026.pdf`
