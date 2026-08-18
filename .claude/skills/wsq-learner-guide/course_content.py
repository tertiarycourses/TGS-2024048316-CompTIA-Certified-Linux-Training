"""course_content.py — SINGLE SOURCE OF TRUTH for TGS-2024048316.

CompTIA Certified Linux+ Training (XK0-006 V8), delivered by Tertiary Infotech
Academy Pte Ltd. Every artifact — the PPT deck, the Lesson Plan (LP), the Learner
Guide (LG DOCX) and its Markdown mirror — is generated from the data below, so all
four stay 100% aligned to the 30 hands-on labs and the CompTIA exam domains.

The 30 labs (goal / build / concepts / step-by-step / test) are loaded from the
batch JSON files produced from the labs/ folder, keyed to the exam objectives.
"""
import html
import json
import os

def _find_repo():
    """Locate the course repo root (the dir containing labs/ and courseware/),
    so this module works from inside a skill folder or anywhere else."""
    env = os.environ.get("COURSE_REPO")
    if env and os.path.isdir(os.path.join(env, "labs")):
        return env
    d = os.path.dirname(os.path.abspath(__file__))
    for _ in range(8):
        if os.path.isdir(os.path.join(d, "labs")) and os.path.isdir(os.path.join(d, "courseware")):
            return d
        parent = os.path.dirname(d)
        if parent == d:
            break
        d = parent
    return os.getcwd()


REPO = _find_repo()
COURSEWARE = os.path.join(REPO, "courseware")     # generated documents land here (outputs only)

# Build INPUTS live beside this single-source module (inside the wsq-learner-guide skill),
# so courseware/ holds only generated deliverables.
_SKILL_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(_SKILL_DIR, "assets")                       # logos
# Structured lab content (extracted from labs/) — the single source of truth.
BATCH_DIR = os.environ.get("LABS_BATCH_DIR", os.path.join(_SKILL_DIR, "data"))

# ---------------------------------------------------------------- course metadata
COURSE = dict(
    title="CompTIA Certified Linux+ Training (XK0-006)",
    short="CompTIA Certified Linux+ Training",
    code="TGS-2024048316",
    exam="XK0-006 V8",
    org="Tertiary Infotech Academy Pte Ltd",
    uen="201200696W",
    version="v3",
    version_date="18 August 2026",
    trainer="Dr. Alfred Ang",
    days=2,
    lms="https://lms-tms.tertiaryinfotech.com/",
    killercoda="https://killercoda.com/playgrounds/scenario/ubuntu",
    repo="https://github.com/tertiarycourses/TGS-2024048316-CompTIA-Certified-Linux-Training",
    register="https://www.tertiarycourses.com.sg/wsq-comptia-certified-linux-training.html",
    practice_exam="https://exams.tertiaryinfotech.com/practice-exams/comptia/comptia-linux-plus",
    exam_voucher="https://store.comptia.org",
    exam_pearson="https://home.pearsonvue.com/comptia",
    org_logo=os.path.join(ASSETS, "tertiary-infotech-logo.png"),
    course_logo=os.path.join(ASSETS, "comptia-linux-logo.png"),
)

# ---------------------------------------------------------------- the five domains
# weight = official XK0-006 exam percentage; objs = the sub-objectives (from the
# official Exam Objectives PDF) so the courseware maps 1:1 to the blueprint.
DOMAINS = [
    dict(num=1, title="System Management", weight=23, labs=list(range(1, 8)),
         objs=[
             ("1.1", "Explain basic Linux concepts"),
             ("1.2", "Summarize Linux device management concepts and tools"),
             ("1.3", "Given a scenario, manage storage in a Linux system"),
             ("1.4", "Given a scenario, manage network services and configurations"),
             ("1.5", "Given a scenario, manage a Linux system using common shell operations"),
             ("1.6", "Given a scenario, perform backup and restore operations"),
             ("1.7", "Summarize virtualization on Linux systems"),
         ]),
    dict(num=2, title="Services and User Management", weight=20, labs=list(range(8, 14)),
         objs=[
             ("2.1", "Given a scenario, manage files and directories"),
             ("2.2", "Given a scenario, perform local account management"),
             ("2.3", "Given a scenario, manage processes and jobs"),
             ("2.4", "Given a scenario, configure and manage software"),
             ("2.5", "Given a scenario, manage Linux using systemd"),
             ("2.6", "Given a scenario, manage applications in a container"),
         ]),
    dict(num=3, title="Security", weight=18, labs=list(range(14, 20)),
         objs=[
             ("3.1", "Summarize authorization, authentication, and accounting methods"),
             ("3.2", "Given a scenario, configure and implement firewalls"),
             ("3.3", "Given a scenario, apply OS hardening techniques"),
             ("3.4", "Explain account hardening techniques and best practices"),
             ("3.5", "Explain cryptographic concepts and technologies"),
             ("3.6", "Explain the importance of compliance and audit procedures"),
         ]),
    dict(num=4, title="Automation, Orchestration, and Scripting", weight=17, labs=list(range(20, 25)),
         objs=[
             ("4.1", "Summarize automation and orchestration use cases and techniques"),
             ("4.2", "Given a scenario, perform automated tasks using shell scripting"),
             ("4.3", "Summarize Python basics used for Linux system administration"),
             ("4.4", "Given a scenario, implement version control using Git"),
             ("4.5", "Summarize best practices and responsible uses of AI"),
         ]),
    dict(num=5, title="Troubleshooting", weight=22, labs=list(range(25, 31)),
         objs=[
             ("5.1", "Summarize monitoring concepts and configurations"),
             ("5.2", "Given a scenario, troubleshoot hardware, storage, and OS issues"),
             ("5.3", "Given a scenario, troubleshoot networking issues"),
             ("5.4", "Given a scenario, troubleshoot security issues"),
             ("5.5", "Given a scenario, troubleshoot performance issues"),
         ]),
]

# Course-level learning outcomes (WSQ house wording).
LEARNING_OUTCOMES = [
    "Explain core Linux concepts — the boot process, the Filesystem Hierarchy Standard, "
    "device management and virtualization — and manage storage, networking and shell operations.",
    "Manage files, local user and group accounts, processes and jobs, software packages, "
    "systemd services and containerized applications on a Linux server.",
    "Apply Linux security: authentication/authorization (sudo, PAM), firewalls, OS and account "
    "hardening, cryptography and compliance/audit procedures.",
    "Automate administration with Ansible, Bash and Python, apply Git version control, and use "
    "AI assistants responsibly and securely.",
    "Monitor a Linux system and analyze and troubleshoot hardware, storage, network, security "
    "and performance issues using the right diagnostic tools.",
]

# Short outcome label per domain (for the visual Learning-Outcome / What-You-Achieved tiles).
DOMAIN_OUTCOMES = {
    1: ("System Management", "Boot & FHS, kernel modules, LVM storage, networking, shell, backup, virtualization."),
    2: ("Services & User Management", "Files, accounts, processes, packages, systemd services, containers."),
    3: ("Security", "AAA/sudo/PAM, firewalls, OS & account hardening, cryptography, compliance & audit."),
    4: ("Automation & Scripting", "Ansible IaC, Bash, Python, Git version control, responsible AI use."),
    5: ("Troubleshooting", "Monitoring, and storage / network / security / performance triage."),
}

# 4 key-concept bullets per domain — drive the visual "Key Concepts" tile grid on the deck
# and the domain intros in the Learner Guide, so the concepts stay aligned everywhere.
DOMAIN_CONCEPTS = {
    1: [
        "The boot chain: firmware/UEFI → GRUB2 → kernel + initramfs → systemd; the FHS defines where files live.",
        "Storage stacks up: partitions → LVM (PV → VG → LV) → filesystem (ext4/xfs/btrfs), mounted via /etc/fstab.",
        "The network stack is read top-down: NetworkManager/Netplan config → ip/ss → DNS resolution → diagnostics.",
        "Shell operations (redirection, pipes, text tools), backup (tar/rsync/dd) and virtualization (QEMU/KVM) round out the domain.",
    ],
    2: [
        "Everything is a file: manage files/links/devices, and local accounts via /etc/passwd, /etc/shadow and /etc/group.",
        "Processes have states, priorities and signals; jobs are scheduled with at, cron and anacron.",
        "Software is installed with apt/dpkg and dnf/rpm, plus language (pip/npm/cargo) and sandboxed (snap/flatpak) managers.",
        "systemd manages services, timers, mounts and targets; containers (Docker/Podman) package and isolate applications.",
    ],
    3: [
        "AAA: authenticate (PAM/SSSD), authorize (sudo, Polkit) and account (journald, auditd) for every action.",
        "Firewalls filter at every layer — ufw → firewalld → nftables/iptables — with stateful inspection and NAT.",
        "OS & account hardening: permissions/ACLs, SELinux, SSH hardening, fail2ban, password policy and MFA.",
        "Cryptography (GPG, LUKS2, OpenSSL, WireGuard) and compliance/audit (AIDE, rkhunter, OpenSCAP, CIS) protect data.",
    ],
    4: [
        "Infrastructure as Code (Ansible, Puppet, OpenTofu) makes configuration repeatable and idempotent.",
        "Bash scripting automates tasks with variables, conditionals, loops, functions and shellcheck-clean code.",
        "Python (venv, modules, PEP 8) extends automation beyond the shell; Git version-controls all of it.",
        "AI assistants are used responsibly: verify-before-paste, redact secrets, and follow a corporate AI policy.",
    ],
    5: [
        "Monitoring vocabulary (SLA/SLI/SLO) and tools (SNMP, Prometheus, node_exporter) establish a baseline.",
        "Troubleshooting is a loop: baseline → reproduce → identify with the right counter → remediate → verify.",
        "Storage/OS faults (ENOSPC, inode exhaustion, failed units) and network faults (DNS, routing, MTU) each have a signature.",
        "Security faults (SELinux, ACLs, certs, lockout) and performance faults (CPU, memory, I/O) close out the domain.",
    ],
}

DAY_TOPICS = {
    1: [1, 2],   # Day 1: Domains 1 & 2
    2: [3, 4, 5],  # Day 2: Domains 3, 4, 5 + Capstone + Assessment
}

# ---------------------------------------------------------------- legacy reference deck
# The full v3 Learner Guide deck (reference/): 22 chapters of teaching slides whose
# content and graphics are carried into the generated deck, REORDERED onto the
# XK0-006 exam domains. Ranges are 1-based inclusive slide numbers in the old deck.
REFERENCE_DECK = os.path.join(
    REPO, "reference",
    "WSQ - Learner Guide Slides - CompTIA Certified Linux+ Training - v3.pptx")

# (chapter, first_slide, last_slide, domain, chapter title)
REFERENCE_CHAPTERS = [
    (1,    17,   87, 1, "Understanding Linux Fundamentals"),
    (2,    88,  186, 2, "Managing Files and Directories"),
    (3,   187,  221, 1, "Configuring and Managing Storage"),
    (4,   222,  321, 2, "Managing Processes and Services"),
    (5,   322,  413, 1, "Using Network Tools and Configuration Files"),
    (6,   414,  550, 2, "Building and Installing Software"),
    (7,   551,  575, 2, "Managing Software Configurations"),
    (8,   577,  666, 3, "Understanding Linux Security Best Practices"),
    (9,   667,  744, 2, "Implementing Identity Management"),
    (10,  745,  823, 3, "Implementing and Configuring Firewalls"),
    (11,  824,  885, 1, "Using Remote Connectivity for System Management"),
    (12,  886, 1039, 3, "Understanding and Applying Access Controls"),
    (13, 1041, 1221, 4, "Automating Tasks via Shell Scripting"),
    (14, 1222, 1260, 2, "Perform Basic Container Operations"),
    (15, 1261, 1335, 4, "Performing Basic Version Control Using Git"),
    (16, 1336, 1376, 4, "Understanding Infrastructure as Code"),
    (17, 1377, 1429, 4, "Understanding Containers, Cloud, and Orchestration"),
    (18, 1431, 1589, 5, "Analyzing and Troubleshooting Storage Issues"),
    (19, 1590, 1686, 5, "Analyzing and Troubleshooting Network Resource Issues"),
    (20, 1687, 1756, 5, "Analyzing and Troubleshooting CPU and Memory Issues"),
    (21, 1757, 1803, 5, "Analyzing and Troubleshooting User and File Permissions"),
    (22, 1804, 1889, 5, "Analyzing and Troubleshooting Common Problems Using Systemd"),
]

# Order the chapters appear inside each domain (follows the objective order).
DOMAIN_CHAPTERS = {
    1: [1, 3, 5, 11],
    2: [2, 9, 4, 6, 7, 14],
    3: [8, 10, 12],
    4: [16, 17, 13, 15],
    5: [18, 19, 20, 21, 22],
}

# ---------------------------------------------------------------- exam objectives (XK0-006 V8)
# Official per-objective summary — short name + one-line description. The deck follows
# these objective-by-objective inside each domain.
OBJ_INFO = {
    "1.1": ("Linux basics", "Identify boot process steps, kernel, filesystems, and architectures."),
    "1.2": ("Device management", "Manage kernel modules, hardware components, and device utilities."),
    "1.3": ("Storage management", "Configure LVM, RAID, partitions, and mounted storage."),
    "1.4": ("Network configuration", "Set up hosts, DNS, interfaces, and network tools."),
    "1.5": ("Shell operations", "Use navigation, editing, redirection, and environment variables."),
    "1.6": ("Backups and restores", "Perform archiving, compression, and data recovery."),
    "1.7": ("Virtualization", "Deploy hypervisors, create VMs, and manage disk images."),
    "2.1": ("Files & directories", "Control permissions, create links, and manage special files."),
    "2.2": ("Account management", "Add, remove, and modify users and groups."),
    "2.3": ("Process control", "Monitor states, adjust priorities, and schedule jobs."),
    "2.4": ("Software management", "Install, update, or remove packages and repositories."),
    "2.5": ("Systems management", "Start, stop, and review services, logs, and timers."),
    "2.6": ("Containers", "Operate container runtimes, manage images, and create networks."),
    "3.1": ("Auth & accounting", "Configure PAM, LDAP, Kerberos, and enable auditing."),
    "3.2": ("Firewalls", "Set firewalls using iptables, nftables, UFW, and zone rules."),
    "3.3": ("OS hardening", "Apply permissions, configure sudo, and secure remote access."),
    "3.4": ("Account security", "Enforce password policies, restrict shells, and enable MFA."),
    "3.5": ("Cryptography", "Encrypt files, use hashing, and manage certificates."),
    "3.6": ("Compliance", "Verify integrity, run scans, and maintain standards."),
    "4.1": ("Automation", "Automate tasks with Ansible, Puppet, and CI/CD tools."),
    "4.2": ("Shell scripting", "Write and troubleshoot variables, functions, and logic flows."),
    "4.3": ("Python basics", "Develop scripts using environments, packages, and data types."),
    "4.4": ("Version control", "Manage code with Git workflows and tagging."),
    "4.5": ("AI best practices", "Apply code generation and prompt engineering responsibly."),
    "5.1": ("System monitoring", "Monitor health, logs, and event alerts."),
    "5.2": ("Hardware/storage", "Diagnose boot, mount, and repair issues."),
    "5.3": ("Networking", "Resolve firewall, routing, DNS, and connectivity problems."),
    "5.4": ("Security", "Fix SELinux, permission, and vulnerability issues."),
    "5.5": ("Performance", "Analyze CPU, memory, I/O, and optimize response times."),
}

# Which legacy chapter's condensed content teaches which objective.
CHAPTER_OBJECTIVE = {
    1: "1.1", 3: "1.3", 5: "1.4", 11: "1.4",
    2: "2.1", 9: "2.2", 4: "2.3", 6: "2.4", 7: "2.4", 14: "2.6",
    12: "3.1", 10: "3.2", 8: "3.3",
    16: "4.1", 17: "4.1", 13: "4.2", 15: "4.4",
    18: "5.2", 22: "5.2", 19: "5.3", 21: "5.4", 20: "5.5",
}

# Old-deck slides NOT carried over: the superseded admin block (1-16), the four
# Part dividers, the retired lab-pointer slides (they reference a lab set that no
# longer exists), and the old summary/close (replaced by the house tail slides).
REFERENCE_DROP = set(range(1, 17)) | {576, 1040, 1430} | set(range(1890, 1897)) | {
    24, 48, 136, 186, 200, 400, 402, 411, 749, 823, 889,
    1039, 1046, 1352, 1399, 1686, 1887,
}


def _unescape(obj):
    if isinstance(obj, str):
        return html.unescape(obj)
    if isinstance(obj, list):
        return [_unescape(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _unescape(v) for k, v in obj.items()}
    return obj


def load_labs():
    labs = []
    for i in range(1, 6):
        path = os.path.join(BATCH_DIR, f"batch{i}.json")
        with open(path) as fh:
            labs.extend(json.load(fh))
    labs = _unescape(labs)
    labs.sort(key=lambda x: x["num"])
    assert len(labs) == 30, f"expected 30 labs, got {len(labs)}"
    return labs


LABS = load_labs()
LABS_BY_NUM = {l["num"]: l for l in LABS}


def domain_labs(dnum):
    return [LABS_BY_NUM[n] for n in DOMAINS[dnum - 1]["labs"]]


if __name__ == "__main__":
    print(f"Loaded {len(LABS)} labs across {len(DOMAINS)} domains")
    for d in DOMAINS:
        print(f"  Domain {d['num']} ({d['weight']}%) {d['title']}: labs {d['labs']}")
    for l in LABS:
        print(f"  Lab {l['num']:>2} [{l['objective']:>8}] {l['title']} — {len(l['steps'])} steps")
