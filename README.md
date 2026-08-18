<div align="center">

# 🐧 CompTIA Certified Linux+ Training

**Master Linux system administration hands-on — and walk out ready for the CompTIA Linux+ XK0-006 exam.**

[![Register](https://img.shields.io/badge/Register-WSQ%20Funded%20Course-C8102E?style=for-the-badge)](https://www.tertiarycourses.com.sg/wsq-comptia-certified-linux-training.html)
[![WSQ](https://img.shields.io/badge/WSQ-TGS--2024048316-1F6FEB)](https://www.tertiarycourses.com.sg/wsq-comptia-certified-linux-training.html)
[![Exam](https://img.shields.io/badge/Exam-XK0--006%20V8-red)](https://www.comptia.org/certifications/linux)
[![Duration](https://img.shields.io/badge/Duration-2%20Days%20%C2%B7%2016%20Hours-10B981)](#course-details)
[![Labs](https://img.shields.io/badge/Hands--On%20Labs-30-7C3AED)](#course-topics--labs)
[![Assessment](https://img.shields.io/badge/Assessment-WA%20%2B%20PP-F59E0B)](#course-details)

**[📝 Register for this course →](https://www.tertiarycourses.com.sg/wsq-comptia-certified-linux-training.html)**

</div>

---

## About This Course

This WSQ-funded 2-day course teaches you to configure, manage, secure, automate and troubleshoot Linux servers, mapped 1:1 to the five **CompTIA Linux+ XK0-006 (V8)** exam domains. It is built for system administrators, DevOps engineers, cloud engineers and IT professionals who want job-ready Linux skills plus a clear path to the Linux+ certification.

Every one of the **30 step-by-step labs** runs free in the browser on the [Killercoda Ubuntu Playground](https://killercoda.com/playgrounds/scenario/ubuntu) — no local install, no virtual machine, no credit card required.

## Learning Outcomes

| # | Learning Outcome |
|---|---|
| LO1 | Develop Linux training guides and evaluate change requests to identify valid and feasible improvements. |
| LO2 | Resolve Linux application issues by developing support guides and analyzing application logs. |
| LO3 | Identify underlying performance issues by analyzing Linux application logs and user feedback. |
| LO4 | Conduct Linux installation and maintenance procedures and propose application enhancements. |

## Course Topics & Labs

The course follows the XK0-006 exam blueprint. Each lab has its own folder under [labs/](labs/) with a full step-by-step `README.md`.

| Exam Domain | Weight | Hands-On Labs |
|---|---|---|
| **1 · System Management** | 23% | [Lab 1 — Boot & FHS](labs/lab-01-boot-fhs/) · [Lab 2 — Kernel & Devices](labs/lab-02-kernel-devices/) · [Lab 3 — Storage & LVM](labs/lab-03-storage-lvm/) · [Lab 4 — Networking](labs/lab-04-networking/) · [Lab 5 — Shell & Text](labs/lab-05-shell/) · [Lab 6 — Backup & Restore](labs/lab-06-backup-restore/) · [Lab 7 — Virtualization](labs/lab-07-virtualization/) |
| **2 · Services and User Management** | 20% | [Lab 8 — Files & Directories](labs/lab-08-files-directories/) · [Lab 9 — Accounts & Groups](labs/lab-09-users-groups/) · [Lab 10 — Processes & Jobs](labs/lab-10-processes-jobs/) · [Lab 11 — Packages](labs/lab-11-packages/) · [Lab 12 — systemd](labs/lab-12-systemd/) · [Lab 13 — Containers](labs/lab-13-containers/) |
| **3 · Security** | 18% | [Lab 14 — AAA: sudo, PAM, Polkit](labs/lab-14-aaa-sudo-pam/) · [Lab 15 — Firewalls](labs/lab-15-firewall/) · [Lab 16 — OS Hardening](labs/lab-16-hardening/) · [Lab 17 — Account Hardening](labs/lab-17-account-hardening/) · [Lab 18 — Cryptography](labs/lab-18-crypto/) · [Lab 19 — Compliance & Audit](labs/lab-19-compliance-audit/) |
| **4 · Automation, Orchestration & Scripting** | 17% | [Lab 20 — Ansible IaC](labs/lab-20-ansible/) · [Lab 21 — Bash Scripting](labs/lab-21-bash-scripting/) · [Lab 22 — Python for Sysadmin](labs/lab-22-python/) · [Lab 23 — Git](labs/lab-23-git/) · [Lab 24 — Responsible AI Use](labs/lab-24-ai-best-practices/) |
| **5 · Troubleshooting** | 22% | [Lab 25 — Monitoring](labs/lab-25-monitoring/) · [Lab 26 — Storage/OS Triage](labs/lab-26-troubleshoot-storage/) · [Lab 27 — Network Triage](labs/lab-27-troubleshoot-network/) · [Lab 28 — Security Triage](labs/lab-28-troubleshoot-security/) · [Lab 29 — Performance Triage](labs/lab-29-troubleshoot-performance/) · [Lab 30 — Capstone](labs/lab-30-capstone/) |

**How to run a lab:** open the [Killercoda Ubuntu Playground](https://killercoda.com/playgrounds/scenario/ubuntu), pick a lab folder, and follow its `README.md` step by step. Reset the playground between labs that change kernel, firewall or systemd state.

## Tools

All tooling is **100% free** — the bulk runs inside the disposable Killercoda VM via `apt`/`dnf` or open-source binaries.

| Tool | Used In | Link |
|---|---|---|
| Killercoda Ubuntu Playground | Every lab | <https://killercoda.com/playgrounds/scenario/ubuntu> |
| Regex Generator | Lab 21 | <https://alfredang.github.io/regexgenerator/> |
| ShellCheck (online) | Lab 21 | <https://www.shellcheck.net/> |
| CompTIA XK0-006 Exam Objectives (PDF) | Every lab | in this repo |
| Practice Exam — CompTIA Linux+ | Exam prep | <https://exams.tertiaryinfotech.com/practice-exams/comptia/comptia-linux-plus> |

Full tool list with install commands: [labs/tools.md](labs/tools.md). Lab references and further practice: [labs/README.md](labs/README.md).

## Repository Structure

```
.
├── courseware/                  # Generated deliverables (single-source build)
│   ├── PPT-CompTIA-Linux-Plus-XK0-006-v3.pptx   # 525-slide training deck (+ PDF)
│   ├── LP-CompTIA-Linux-Plus-XK0-006-v3.docx    # 2-day Lesson Plan (+ PDF)
│   ├── LG-CompTIA-Linux-Plus-XK0-006-v3.docx    # Learner Guide (+ PDF)
│   ├── slide_map.json           # lab/domain/chapter → deck slide numbers
│   └── archive/                 # superseded versions
├── labs/                        # 30 hands-on labs — one folder per lab
│   ├── lab-01-boot-fhs/ … lab-30-capstone/
│   ├── README.md                # lab index + references
│   └── tools.md                 # free-tool catalogue
├── .claude/                     # WSQ courseware build toolchain (skills + scripts)
└── CompTIA Linux+ XK0-006 V8 Exam Objectives (4.0).pdf
```

> The confidential `assessment/` folder (question papers + answer keys) is **not** in this repository — it is distributed to trainers via Google Drive only.

The deck follows the exam blueprint objective-by-objective (1.1 → 5.5) in the standard house slide format — condensed teaching content with key screenshots, full step-by-step lab slides, and the exam-domain blueprint at the start and the CompTIA exam registration + practice exam at the end.

## Course Details

| | |
|---|---|
| **Course Title** | WSQ — CompTIA Certified Linux+ Training |
| **Course Code** | TGS-2024048316 |
| **Certification Exam** | CompTIA Linux+ XK0-006 (V8) |
| **Duration** | 2 days · 16 hours (9:30am – 6:30pm) |
| **Assessment** | Written Assessment (SAQ, 1 hr) + Practical Performance (PP, 1 hr) — open book |
| **Mode** | Classroom / synchronous e-learning, hands-on |
| **Provider** | Tertiary Infotech Academy Pte Ltd (UEN 201200696W) |

**Funding:** This course is WSQ-funded (SkillsFuture Singapore). Singaporeans and PRs may be eligible for course-fee funding and may use SkillsFuture Credit — see the [course page](https://www.tertiarycourses.com.sg/wsq-comptia-certified-linux-training.html) for current funding tiers.

## Building the Courseware

All artifacts are generated from a single content module (`.claude/skills/wsq-learner-guide/course_content.py`) so the deck, Lesson Plan, Learner Guide and labs never drift apart. Build the deck first — it writes `slide_map.json`, which the Lesson Plan cites:

```bash
python3 .claude/skills/wsq-slides/build_slides.py            # deck + slide_map.json
python3 .claude/skills/wsq-lesson-plan/build_lesson_plan.py  # Lesson Plan
python3 .claude/skills/wsq-learner-guide/build_learner_guide.py  # Learner Guide
```

## Contact

- 🌐 **Course page & registration:** <https://www.tertiarycourses.com.sg/wsq-comptia-certified-linux-training.html>
- ✉️ **Email:** enquiry@tertiaryinfotech.com
- 📞 **Tel:** +65 6100 0613
- 🎓 **LMS/TMS:** <https://lms-tms.tertiaryinfotech.com/>

---

<div align="center">

**Ready to become Linux+ certified?**

**[📝 Register for the WSQ CompTIA Certified Linux+ Training →](https://www.tertiarycourses.com.sg/wsq-comptia-certified-linux-training.html)**

© 2026 Tertiary Infotech Academy Pte Ltd · UEN 201200696W

</div>
