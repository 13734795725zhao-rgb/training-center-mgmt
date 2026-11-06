<div align="center">
  <img src="docs/assets/banner.svg" alt="Training Center Management System" width="880">
</div>

<br>

<div align="center">

**The admin system a small training centre actually needs — students, attendance, tuition, refunds, reports.**

Runs on one office computer. Your data stays in one file on that computer. No cloud, no monthly fee.

[![Python](https://img.shields.io/badge/python-3.9%2B-3776AB)](#requirements)
[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

[Features](#features) · [Get started](#get-started) · [Who can do what](#who-can-do-what) · [Questions](#common-questions)

</div>

---

> [!CAUTION]
> **Do these three things before entering any real student data.**
>
> **1. Set `FLASK_SECRET_KEY`.** The code falls back to a default that is written into it. If you leave it, every installation shares the same key and a login could be forged. Set any long random string before starting.
>
> **2. Change the starting passwords.** The first launch creates `admin` / `admin123` and `principal` / `principal123`. Those are not passwords to leave on an office network.
>
> **3. Never upload the database or the photos.** `students.db` and `uploads/` hold students' names, parents' phone numbers, payment records and photographs of children. Both are excluded from this repository, and **no data ships with it** — an empty database is created on first start.

---

## Features

| Module | What it does |
|---|---|
| **Student records** | Each student's details, an automatically generated student number, parent contacts, class history and status |
| **Student portfolio** | Awards, projects and milestones per student, editable by staff and teachers |
| **Attendance** | Mark attendance lesson by lesson, per student, rolled up per course |
| **Courses and subjects** | The course list with subject grouping, lesson counts and prices |
| **Teachers** | Teacher records, the subjects they teach and the courses they run |
| **Tuition payments** | Record a payment against a class, attach a photo of the receipt, and see the balance |
| **Refunds** | Request a refund against a paid class, attach proof, keep the approval trail |
| **Lesson hours** | How many lessons each student has left, reduced automatically by attendance |
| **Reports** | Enrolment, income and attendance statistics, filtered by date range |
| **Excel import** | Bring a student list in from a spreadsheet. It checks the file first and tells you which row is wrong |
| **Activity log** | Every change is recorded — who changed it, what they changed, and when |
| **Permissions** | Four roles, checked on the server, not just hidden in the page |
| **Automatic numbering** | Student numbers and receipt numbers generated without duplicates |
| **Works on a phone** | Pages adapt to a phone screen, for the front desk |

---

## Who can do what

| Role | What they can do |
|---|---|
| **Admin** | Everything, including user accounts and system settings |
| **Principal** | All day-to-day work — students, courses, payments, refunds, reports. Cannot delete an admin |
| **Front desk / staff** | Students, attendance, payments, refunds, portfolios, subjects |
| **Teacher** | View students, take attendance, edit portfolios, see their own courses |

Pages a role cannot use reject them even if the address is typed by hand.

---

## Requirements

- Python 3.9 or newer
- Two packages: `flask` and `openpyxl`

## Get started

```bash
git clone https://github.com/13734795725zhao-rgb/training-center-mgmt.git
cd training-center-mgmt

pip install -r requirements.txt

# required: use your own key
export FLASK_SECRET_KEY="a-long-random-string-of-your-own"     # Linux / macOS
# $env:FLASK_SECRET_KEY = "..."                                # PowerShell

python app.py
```

Open **http://127.0.0.1:5000/** and log in. The database is created automatically on first start — you do not need to set anything up.

### For real day-to-day use

`python app.py` uses a development mode that is fine for trying the system out. To let colleagues use it, start it properly:

```bash
pip install waitress
waitress-serve --listen=0.0.0.0:8000 app:app
```

**Backing up is copying files.** `students.db` and the `uploads/` folder are the entire system — copy them to a USB stick or another drive and you have a backup.

---

## Common questions

| Problem | What to do |
|---|---|
| A colleague cannot open it | Allow port 5000 (or 8000) through the firewall on your computer, and have them use your computer's network address |
| The admin password is forgotten | Another admin resets it under **User management**. If all admins are locked out, the database has to be edited directly |
| The Excel import complains | Read the message — it names the row and the column that is wrong. Usually a date or phone number is stored as a number |
| I want to move to another computer | Copy the whole folder, including `students.db` and `uploads/` |
| A student was deleted by mistake | That cannot be undone — their attendance, payments and refunds go with them. Check the activity log to see who did it and when |
| I want to start over | Stop the program, delete `students.db` and `uploads/`, and start it again |

---

## Good to know

1. **The interface is in Chinese**, and so is the data.
2. **One copy per centre.** If you run several branches, give each its own copy rather than sharing one.
3. **Change the passwords and set the key before anyone else touches it.**
4. **Built for a local network.** Putting it on the public internet needs HTTPS and stricter access control first.
5. **Deleting a student deletes everything about them** — including their payment and refund records. Only admin and principal can do it, and it cannot be undone.

---

## More

| Document | Contents |
|---|---|
| [`docs/PRD.zh-CN.md`](docs/PRD.zh-CN.md) | Full feature description, numbering rules, permission model (Chinese) |
| [`docs/TECHNICAL-DESIGN.zh-CN.md`](docs/TECHNICAL-DESIGN.zh-CN.md) | Technical design (Chinese) |

## License

[MIT](LICENSE) © 2026 幻蝶

---

<div align="center">
<sub>Small centres deserve software that fits on a USB stick.</sub>
</div>
