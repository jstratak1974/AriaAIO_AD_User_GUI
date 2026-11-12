#!/usr/bin/env python3
"""
ARIA / Eclipse v18.x Active Directory User Provisioning GUI

Companion workbook: ARIA_AD_User_Template_v18_3.xlsx

Purpose
-------
- Load AD user definitions from Excel.
- Review/edit/add/remove rows before execution.
- Validate ARIA/Varian-oriented naming, OU, group and password rules.
- Use Microsoft's ActiveDirectory PowerShell module to create users and add group memberships.
- Support AD preflight and non-destructive What-If before actual creation.

Important
---------
This tool automates only the Windows Active Directory step. ARIA/Eclipse v18.x users must
subsequently exist/be active in Shared Framework / Varian Service Portal and receive the
appropriate ARIA application rights.

Requirements on the Windows administration PC
----------------------------------------------
- Python 3.x
- openpyxl:  pip install openpyxl
- Windows PowerShell with RSAT ActiveDirectory module
- AD rights sufficient to create users and add group memberships

Safety choices
--------------
- Existing AD users are skipped, never modified.
- Passwords are not included in previews or audit logs.
- Prompt mode is recommended so passwords are not stored in Excel.
- Use What-If before Execute Create.
"""

from __future__ import annotations

import ctypes
import json
import os
import re
import secrets
import string
import subprocess
import sys
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

APP_TITLE = "ARIA v18.x AD User Provisioning"

EXPECTED_COLUMNS = [
    "Enabled", "FirstName", "Initials", "LastName", "DisplayName",
    "SamAccountName", "UserPrincipalName", "OU", "Role", "ADGroups",
    "Department", "Description", "Email", "ChangePasswordAtLogon",
    "PasswordNeverExpires", "PasswordMode", "InitialPassword"
]

BOOL_COLUMNS = {"Enabled", "ChangePasswordAtLogon", "PasswordNeverExpires"}
CHOICES = {
    "Enabled": ["TRUE", "FALSE"],
    "Role": ["Standard User", "Administrator"],
    "ChangePasswordAtLogon": ["TRUE", "FALSE"],
    "PasswordNeverExpires": ["TRUE", "FALSE"],
    "PasswordMode": ["Prompt", "Generate", "Excel"],
}

# Manual restriction used for ordinary ARIA interactive accounts in this tool.
USERNAME_RE = re.compile(r"^[A-Za-z0-9]+$")

POWERSHELL_PRECHECK = r'''
$ErrorActionPreference = "Stop"
Import-Module ActiveDirectory
$payload = [Console]::In.ReadToEnd() | ConvertFrom-Json

$serverArgs = @{}
if ($payload.domainController) {
    $serverArgs["Server"] = [string]$payload.domainController
}

$result = [ordered]@{}
$result.Computer = $env:COMPUTERNAME
$result.WindowsUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$result.Domain = (Get-ADDomain @serverArgs).DNSRoot

$policy = Get-ADDefaultDomainPasswordPolicy @serverArgs
$result.PasswordPolicy = [ordered]@{
    MinPasswordLength = $policy.MinPasswordLength
    LockoutThreshold = $policy.LockoutThreshold
    LockoutDuration = [string]$policy.LockoutDuration
    ComplexityEnabled = $policy.ComplexityEnabled
}

$ou = [string]$payload.defaultOU
if ($ou) {
    Get-ADOrganizationalUnit -Identity $ou @serverArgs -ErrorAction Stop | Out-Null
    $result.OU = "OK: $ou"
}

$groupResult = @()
foreach ($g in @($payload.groups)) {
    if ([string]::IsNullOrWhiteSpace([string]$g)) { continue }
    $obj = Get-ADGroup -Identity ([string]$g) @serverArgs -ErrorAction Stop
    $groupResult += "$($obj.Name) [$($obj.GroupScope)]"
}
$result.Groups = $groupResult

[pscustomobject]$result | ConvertTo-Json -Depth 6 -Compress
'''

POWERSHELL_CREATE = r'''
$ErrorActionPreference = "Stop"
Import-Module ActiveDirectory
$payload = [Console]::In.ReadToEnd() | ConvertFrom-Json

$serverArgs = @{}
if ($payload.domainController) {
    $serverArgs["Server"] = [string]$payload.domainController
}

$results = @()
foreach ($u in @($payload.users)) {
    $r = [ordered]@{
        SamAccountName = [string]$u.sam
        DisplayName = [string]$u.displayName
        Status = ""
        Message = ""
        Groups = (@($u.groups) -join "; ")
    }

    try {
        $sam = [string]$u.sam

        # Validate referenced OU and groups before creating anything.
        Get-ADOrganizationalUnit -Identity ([string]$u.ou) @serverArgs -ErrorAction Stop | Out-Null
        foreach ($g in @($u.groups)) {
            Get-ADGroup -Identity ([string]$g) @serverArgs -ErrorAction Stop | Out-Null
        }

        # Deliberately do not modify existing accounts.
        $existing = Get-ADUser -Filter "SamAccountName -eq '$sam'" @serverArgs -ErrorAction Stop
        if ($existing) {
            $r.Status = "Skipped"
            $r.Message = "AD user already exists; no changes were made."
            $results += [pscustomobject]$r
            continue
        }

        if ([bool]$payload.whatIf) {
            $r.Status = "Planned"
            $r.Message = "OU and groups validated. User does not yet exist."
            $results += [pscustomobject]$r
            continue
        }

        $securePassword = ConvertTo-SecureString ([string]$u.password) -AsPlainText -Force

        $newUser = @{
            Name                  = [string]$u.displayName
            GivenName             = [string]$u.firstName
            Surname               = [string]$u.lastName
            DisplayName           = [string]$u.displayName
            SamAccountName        = $sam
            UserPrincipalName     = [string]$u.upn
            Path                  = [string]$u.ou
            AccountPassword       = $securePassword
            Enabled               = $true
            ChangePasswordAtLogon = [bool]$u.changePasswordAtLogon
            PasswordNeverExpires  = [bool]$u.passwordNeverExpires
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$u.initials)) {
            $newUser["Initials"] = [string]$u.initials
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$u.department)) {
            $newUser["Department"] = [string]$u.department
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$u.description)) {
            $newUser["Description"] = [string]$u.description
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$u.email)) {
            $newUser["EmailAddress"] = [string]$u.email
        }

        New-ADUser @newUser @serverArgs -ErrorAction Stop

        foreach ($g in @($u.groups)) {
            Add-ADGroupMember -Identity ([string]$g) -Members $sam @serverArgs -ErrorAction Stop
        }

        $r.Status = "Created"
        $r.Message = "AD account created and group memberships applied."
    }
    catch {
        $r.Status = "Error"
        $r.Message = $_.Exception.Message
    }

    $results += [pscustomobject]$r
}

[pscustomobject]@{
    WhatIf = [bool]$payload.whatIf
    Results = $results
} | ConvertTo-Json -Depth 8 -Compress
'''


def as_bool(value, default=False):
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() in {"true", "1", "yes", "y", "on"}


def bool_text(value):
    return "TRUE" if as_bool(value) else "FALSE"


def password_categories(password: str) -> int:
    return sum([
        any(c.isupper() for c in password),
        any(c.islower() for c in password),
        any(c.isdigit() for c in password),
        any(not c.isalnum() for c in password),
    ])


def validate_interactive_password(password: str):
    if len(password) < 12:
        return False, "Password must be at least 12 characters for this ARIA-oriented workflow."
    if password_categories(password) < 3:
        return False, "Password must contain at least 3 categories: uppercase, lowercase, number, special."
    return True, ""


def generate_password(length=16):
    # Avoid special characters that can be troublesome in some Varian service-account contexts.
    special = "!#$%&*+-=?_"
    alphabet = string.ascii_letters + string.digits + special
    while True:
        pwd = "".join(secrets.choice(alphabet) for _ in range(length))
        if password_categories(pwd) >= 4:
            return pwd


def split_groups(text):
    return [g.strip() for g in str(text or "").split(";") if g.strip()]


def run_powershell(script, payload):
    if os.name != "nt":
        raise RuntimeError("AD execution is supported only on Windows.")

    cmd = [
        "powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive",
        "-ExecutionPolicy", "Bypass", "-Command", script
    ]
    proc = subprocess.run(
        cmd,
        input=json.dumps(payload),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if proc.returncode != 0:
        msg = proc.stderr.strip() or proc.stdout.strip() or f"PowerShell exited with code {proc.returncode}"
        raise RuntimeError(msg)

    output = proc.stdout.strip()
    if not output:
        return {}
    try:
        return json.loads(output)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"PowerShell returned unexpected output:\n{output}") from exc


def is_local_admin():
    if os.name != "nt":
        return False
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1580x850")
        self.minsize(1180, 650)

        self.xlsx_path = tk.StringVar()
        self.rows = []
        self.editor = None

        self.domain_var = tk.StringVar(value="oncology.local")
        self.dc_var = tk.StringVar(value="")
        self.ou_var = tk.StringVar(value="OU=Varian,DC=oncology,DC=local")
        self.upn_var = tk.StringVar(value="oncology.local")
        self.standard_group_var = tk.StringVar(value="Varian Users")
        self.admin_groups_var = tk.StringVar(value="Varian Users;Varian Administrators")

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        ttk.Label(top, text="Excel file:").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.xlsx_path, width=90).grid(row=0, column=1, columnspan=5, sticky="ew", padx=4)
        ttk.Button(top, text="Load Excel", command=self.load_excel_dialog).grid(row=0, column=6, padx=4)
        ttk.Button(top, text="Save Edited Excel", command=self.save_excel_dialog).grid(row=0, column=7, padx=4)

        labels = [
            ("Domain FQDN", self.domain_var),
            ("Domain Controller", self.dc_var),
            ("Default OU", self.ou_var),
            ("UPN Suffix", self.upn_var),
            ("Standard group", self.standard_group_var),
            ("Admin groups", self.admin_groups_var),
        ]
        for idx, (label, var) in enumerate(labels):
            row = 1 + idx // 3
            col = (idx % 3) * 3
            ttk.Label(top, text=label + ":").grid(row=row, column=col, sticky="w", pady=3)
            ttk.Entry(top, textvariable=var, width=45).grid(row=row, column=col + 1, columnspan=2, sticky="ew", padx=4)

        for col in range(8):
            top.columnconfigure(col, weight=1 if col in (1, 2, 4, 5) else 0)

        ttk.Label(
            self,
            text=(
                "Workflow: Excel → review/edit → Validate → AD Preflight / What-If → Execute. "
                "AD creation does not assign Shared Framework / VSP application rights."
            ),
            padding=(10, 2),
        ).pack(fill="x")

        toolbar = ttk.Frame(self, padding=(8, 4))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="Add Row", command=self.add_row).pack(side="left", padx=3)
        ttk.Button(toolbar, text="Remove Selected", command=self.remove_selected).pack(side="left", padx=3)
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(toolbar, text="Validate", command=self.validate_gui).pack(side="left", padx=3)
        ttk.Button(toolbar, text="AD Preflight", command=self.ad_preflight).pack(side="left", padx=3)
        ttk.Button(toolbar, text="Preview Commands", command=self.preview_commands).pack(side="left", padx=3)
        ttk.Button(toolbar, text="What-If", command=lambda: self.execute(True)).pack(side="left", padx=3)
        ttk.Button(toolbar, text="EXECUTE CREATE", command=lambda: self.execute(False)).pack(side="left", padx=12)

        frame = ttk.Frame(self, padding=8)
        frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(frame, columns=EXPECTED_COLUMNS, show="headings", selectmode="extended")
        xbar = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        ybar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        widths = {
            "Enabled": 80, "FirstName": 110, "Initials": 75, "LastName": 110,
            "DisplayName": 165, "SamAccountName": 140, "UserPrincipalName": 210,
            "OU": 260, "Role": 130, "ADGroups": 230, "Department": 160,
            "Description": 240, "Email": 200, "ChangePasswordAtLogon": 165,
            "PasswordNeverExpires": 165, "PasswordMode": 120, "InitialPassword": 140,
        }
        for col in EXPECTED_COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=widths.get(col, 120), minwidth=70, anchor="w")

        self.tree.bind("<Double-1>", self.start_edit)

        bottom = ttk.Frame(self, padding=8)
        bottom.pack(fill="x")
        ttk.Label(bottom, text="Log / results:").pack(anchor="w")
        self.log = tk.Text(bottom, height=9, wrap="word")
        self.log.pack(fill="x")
        self.log.insert("end", "Load the supplied Excel template to begin.\n")
        self.log.configure(state="disabled")

    def write_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def load_excel_dialog(self):
        if load_workbook is None:
            messagebox.showerror("Missing dependency", "Install openpyxl first:\n\npip install openpyxl")
            return
        path = filedialog.askopenfilename(filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.load_excel(path)

    def load_excel(self, path):
        try:
            wb = load_workbook(path, data_only=False)
            if "Users" not in wb.sheetnames:
                raise ValueError("Workbook must contain a 'Users' sheet.")

            if "Config" in wb.sheetnames:
                ws = wb["Config"]
                cfg = {}
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row[0] is not None:
                        cfg[str(row[0]).strip()] = row[1]
                self.domain_var.set(str(cfg.get("DomainFQDN") or self.domain_var.get()))
                self.dc_var.set(str(cfg.get("DomainController") or ""))
                self.ou_var.set(str(cfg.get("DefaultOU") or self.ou_var.get()))
                self.upn_var.set(str(cfg.get("UPNSuffix") or self.upn_var.get()))
                self.standard_group_var.set(str(cfg.get("StandardGroup") or self.standard_group_var.get()))
                self.admin_groups_var.set(str(cfg.get("AdminGroups") or self.admin_groups_var.get()))

            ws = wb["Users"]
            header_values = [c.value for c in ws[1]]
            missing = [c for c in EXPECTED_COLUMNS if c not in header_values]
            if missing:
                raise ValueError("Missing Users columns: " + ", ".join(missing))

            colmap = {name: header_values.index(name) for name in EXPECTED_COLUMNS}
            data = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not any(v not in (None, "") for v in row):
                    continue
                item = {}
                for col in EXPECTED_COLUMNS:
                    value = row[colmap[col]]
                    item[col] = bool_text(value) if col in BOOL_COLUMNS else ("" if value is None else str(value))
                data.append(item)

            self.rows = data
            self.xlsx_path.set(path)
            self.refresh_tree()
            self.write_log(f"Loaded {len(data)} row(s) from {path}")
        except Exception as exc:
            messagebox.showerror("Load failed", str(exc))

    def refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, row in enumerate(self.rows):
            self.tree.insert("", "end", iid=str(idx), values=[row.get(c, "") for c in EXPECTED_COLUMNS])

    def start_edit(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return

        col_index = int(col_id[1:]) - 1
        column = EXPECTED_COLUMNS[col_index]
        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, width, height = bbox
        current = self.tree.set(row_id, column)

        if self.editor is not None:
            try:
                self.editor.destroy()
            except Exception:
                pass

        if column in CHOICES:
            editor = ttk.Combobox(self.tree, values=CHOICES[column], state="readonly")
            editor.set(current)
        else:
            editor = ttk.Entry(self.tree, show="*" if column == "InitialPassword" else None)
            editor.insert(0, current)

        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()
        self.editor = editor

        def commit(_event=None):
            value = editor.get()
            self.tree.set(row_id, column, value)
            self.rows[int(row_id)][column] = value
            editor.destroy()
            self.editor = None

        editor.bind("<Return>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind("<Escape>", lambda _e: editor.destroy())

    def add_row(self):
        row = {c: "" for c in EXPECTED_COLUMNS}
        row.update({
            "Enabled": "TRUE",
            "Role": "Standard User",
            "ADGroups": self.standard_group_var.get(),
            "OU": self.ou_var.get(),
            "ChangePasswordAtLogon": "TRUE",
            "PasswordNeverExpires": "FALSE",
            "PasswordMode": "Prompt",
        })
        self.rows.append(row)
        self.refresh_tree()
        self.write_log("Added blank user row.")

    def remove_selected(self):
        selected = sorted([int(i) for i in self.tree.selection()], reverse=True)
        if not selected:
            return
        for idx in selected:
            if 0 <= idx < len(self.rows):
                self.rows.pop(idx)
        self.refresh_tree()
        self.write_log(f"Removed {len(selected)} selected row(s).")

    def normalized_users(self):
        result = []
        for excel_row, raw in enumerate(self.rows, start=2):
            if not as_bool(raw.get("Enabled"), True):
                continue

            first = raw.get("FirstName", "").strip()
            initials = raw.get("Initials", "").strip()
            last = raw.get("LastName", "").strip()
            display = raw.get("DisplayName", "").strip()
            if not display:
                display = " ".join(p for p in [first, (initials + ".") if initials else "", last] if p).strip()

            sam = raw.get("SamAccountName", "").strip()
            upn = raw.get("UserPrincipalName", "").strip()
            if not upn and sam:
                upn = f"{sam}@{self.upn_var.get().strip()}"

            ou = raw.get("OU", "").strip() or self.ou_var.get().strip()
            role = raw.get("Role", "").strip() or "Standard User"
            groups = split_groups(raw.get("ADGroups", ""))
            if not groups:
                groups = split_groups(self.admin_groups_var.get() if role == "Administrator" else self.standard_group_var.get())

            result.append({
                "excelRow": excel_row,
                "firstName": first,
                "initials": initials,
                "lastName": last,
                "displayName": display,
                "sam": sam,
                "upn": upn,
                "ou": ou,
                "role": role,
                "groups": groups,
                "department": raw.get("Department", "").strip(),
                "description": raw.get("Description", "").strip(),
                "email": raw.get("Email", "").strip(),
                "changePasswordAtLogon": as_bool(raw.get("ChangePasswordAtLogon"), True),
                "passwordNeverExpires": as_bool(raw.get("PasswordNeverExpires"), False),
                "passwordMode": raw.get("PasswordMode", "Prompt").strip() or "Prompt",
                "initialPassword": raw.get("InitialPassword", ""),
            })
        return result

    def validate_users(self):
        errors, warnings = [], []
        users = self.normalized_users()
        seen = set()

        for u in users:
            prefix = f"Excel row {u['excelRow']} ({u['sam'] or 'no username'})"
            if not u["firstName"]:
                errors.append(f"{prefix}: FirstName is required.")
            if not u["lastName"]:
                errors.append(f"{prefix}: LastName is required.")
            if not u["sam"]:
                errors.append(f"{prefix}: SamAccountName is required.")

            sam = u["sam"]
            if sam:
                if len(sam) > 16:
                    errors.append(f"{prefix}: SamAccountName exceeds 16 characters.")
                if not USERNAME_RE.fullmatch(sam):
                    errors.append(f"{prefix}: SamAccountName must contain only letters and numbers.")
                key = sam.lower()
                if key in seen:
                    errors.append(f"{prefix}: duplicate SamAccountName in workbook.")
                seen.add(key)

            if not u["upn"] or "@" not in u["upn"]:
                errors.append(f"{prefix}: UserPrincipalName must be a user@domain value.")
            if not u["ou"]:
                errors.append(f"{prefix}: OU is required.")
            if u["role"] not in {"Standard User", "Administrator"}:
                errors.append(f"{prefix}: Role must be Standard User or Administrator.")
            if not u["groups"]:
                errors.append(f"{prefix}: at least one AD group is required.")
            if u["passwordNeverExpires"] and u["changePasswordAtLogon"]:
                errors.append(f"{prefix}: PasswordNeverExpires and ChangePasswordAtLogon cannot both be TRUE.")
            if u["passwordNeverExpires"]:
                warnings.append(f"{prefix}: PasswordNeverExpires is TRUE; this is normally for service/background logins, not ordinary interactive users.")
            if u["passwordMode"] not in {"Prompt", "Generate", "Excel"}:
                errors.append(f"{prefix}: PasswordMode must be Prompt, Generate, or Excel.")
            if u["passwordMode"] == "Excel":
                ok, msg = validate_interactive_password(str(u["initialPassword"] or ""))
                if not ok:
                    errors.append(f"{prefix}: Excel password: {msg}")
                warnings.append(f"{prefix}: InitialPassword is stored in plaintext in Excel.")

        if not users:
            errors.append("No enabled users are available for creation.")
        return users, errors, warnings

    def validate_gui(self):
        users, errors, warnings = self.validate_users()
        self.write_log(f"Validation: {len(users)} enabled user(s), {len(errors)} error(s), {len(warnings)} warning(s).")
        if errors:
            messagebox.showerror("Validation errors", "\n".join(errors[:25]))
            return False
        if warnings:
            messagebox.showwarning("Validation warnings", "\n".join(warnings[:25]))
        else:
            messagebox.showinfo("Validation", "Validation passed.")
        return True

    def ad_preflight(self):
        _, errors, _ = self.validate_users()
        if errors:
            messagebox.showerror("Validation errors", "\n".join(errors[:25]))
            return

        groups = set(split_groups(self.standard_group_var.get()) + split_groups(self.admin_groups_var.get()))
        for u in self.normalized_users():
            groups.update(u["groups"])

        payload = {
            "domainController": self.dc_var.get().strip(),
            "defaultOU": self.ou_var.get().strip(),
            "groups": sorted(groups),
        }
        try:
            result = run_powershell(POWERSHELL_PRECHECK, payload)
            policy = result.get("PasswordPolicy") or {}
            lines = [
                f"Windows identity: {result.get('WindowsUser', '')}",
                f"Domain: {result.get('Domain', '')}",
                f"OU: {result.get('OU', '')}",
                "Groups: " + ", ".join(result.get("Groups", []) or []),
                (
                    "Domain policy: "
                    f"MinPasswordLength={policy.get('MinPasswordLength')}, "
                    f"LockoutThreshold={policy.get('LockoutThreshold')}, "
                    f"LockoutDuration={policy.get('LockoutDuration')}, "
                    f"ComplexityEnabled={policy.get('ComplexityEnabled')}"
                ),
            ]
            if os.name == "nt" and not is_local_admin():
                lines.append("Note: Python is not elevated locally; delegated AD rights may still be sufficient.")
            msg = "\n".join(lines)
            self.write_log("AD preflight passed. " + msg.replace("\n", " | "))
            messagebox.showinfo("AD Preflight passed", msg)
        except Exception as exc:
            self.write_log(f"AD preflight FAILED: {exc}")
            messagebox.showerror("AD Preflight failed", str(exc))

    def preview_commands(self):
        users, errors, _ = self.validate_users()
        if errors:
            messagebox.showerror("Validation errors", "\n".join(errors[:25]))
            return

        lines = []
        for u in users:
            pwd_desc = {
                "Prompt": "<prompt securely at execution>",
                "Generate": "<generated temporary password>",
                "Excel": "<password from workbook; hidden>",
            }[u["passwordMode"]]
            lines.append(
                f"New-ADUser -Name {json.dumps(u['displayName'])} "
                f"-SamAccountName {u['sam']} -UserPrincipalName {u['upn']} "
                f"-Path {json.dumps(u['ou'])} -AccountPassword {pwd_desc}"
            )
            for group in u["groups"]:
                lines.append(f"Add-ADGroupMember -Identity {json.dumps(group)} -Members {u['sam']}")
            lines.append("")
        self.show_text_window("Command preview (passwords hidden)", "\n".join(lines))

    def resolve_passwords(self, users):
        generated = {}
        for u in users:
            mode = u["passwordMode"]
            if mode == "Generate":
                password = generate_password()
                generated[u["sam"]] = password
            elif mode == "Excel":
                password = str(u["initialPassword"] or "")
            else:
                p1 = simpledialog.askstring(
                    "Initial password",
                    f"Enter initial AD password for {u['displayName']} ({u['sam']}):",
                    show="*", parent=self,
                )
                if p1 is None:
                    raise RuntimeError("Password entry cancelled.")
                p2 = simpledialog.askstring(
                    "Confirm password",
                    f"Confirm password for {u['sam']}:",
                    show="*", parent=self,
                )
                if p2 is None or p1 != p2:
                    raise RuntimeError(f"Password confirmation failed for {u['sam']}.")
                password = p1

            ok, msg = validate_interactive_password(password)
            if not ok:
                raise RuntimeError(f"{u['sam']}: {msg}")
            u["password"] = password
        return generated

    def execute(self, what_if):
        users, errors, _ = self.validate_users()
        if errors:
            messagebox.showerror("Validation errors", "\n".join(errors[:25]))
            return

        generated = {}
        if not what_if:
            if not messagebox.askyesno(
                "Confirm AD creation",
                f"This will CREATE {len(users)} enabled Active Directory account(s) and add group memberships.\n\n"
                "Existing accounts are skipped and never modified.\n\nContinue?"
            ):
                return
            try:
                generated = self.resolve_passwords(users)
            except Exception as exc:
                messagebox.showerror("Password error", str(exc))
                return

        payload_users = []
        for u in users:
            payload_users.append({
                "firstName": u["firstName"],
                "initials": u["initials"],
                "lastName": u["lastName"],
                "displayName": u["displayName"],
                "sam": u["sam"],
                "upn": u["upn"],
                "ou": u["ou"],
                "groups": u["groups"],
                "department": u["department"],
                "description": u["description"],
                "email": u["email"],
                "changePasswordAtLogon": u["changePasswordAtLogon"],
                "passwordNeverExpires": u["passwordNeverExpires"],
                "password": "" if what_if else u["password"],
            })

        payload = {
            "domainController": self.dc_var.get().strip(),
            "whatIf": what_if,
            "users": payload_users,
        }

        try:
            result = run_powershell(POWERSHELL_CREATE, payload)
            results = result.get("Results", [])
            if isinstance(results, dict):
                results = [results]

            display_lines = []
            for r in results:
                line = f"{r.get('SamAccountName', '')}: {r.get('Status', '')} — {r.get('Message', '')}"
                display_lines.append(line)
                self.write_log(line)

            self.write_audit_log(results, what_if)

            if generated and not what_if:
                lines = [
                    "Generated temporary passwords (shown once; not written to the audit log):",
                    "",
                ]
                for sam, pwd in generated.items():
                    lines.append(f"{sam}\t{pwd}")
                lines += ["", "Transfer these credentials securely and close this window when finished."]
                self.show_text_window("Generated temporary passwords", "\n".join(lines))

            self.show_text_window(
                "What-If complete" if what_if else "AD creation complete",
                "\n".join(display_lines) or "No results returned."
            )
        except Exception as exc:
            self.write_log(f"Execution FAILED: {exc}")
            messagebox.showerror("Execution failed", str(exc))
        finally:
            for u in users:
                u.pop("password", None)

    def write_audit_log(self, results, what_if):
        base = Path(self.xlsx_path.get()).resolve().parent if self.xlsx_path.get() else Path.cwd()
        path = base / "ARIA_AD_User_Creation.log"
        try:
            with path.open("a", encoding="utf-8") as handle:
                for r in results:
                    handle.write(
                        f"{datetime.now().isoformat(timespec='seconds')}\t"
                        f"{'WHATIF' if what_if else 'EXECUTE'}\t"
                        f"{r.get('SamAccountName', '')}\t{r.get('Status', '')}\t"
                        f"{r.get('Groups', '')}\t{r.get('Message', '')}\n"
                    )
            self.write_log(f"Audit log: {path}")
        except Exception as exc:
            self.write_log(f"Could not write audit log: {exc}")

    def show_text_window(self, title, content):
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("1000x600")
        text = tk.Text(win, wrap="word")
        text.pack(fill="both", expand=True, padx=8, pady=8)
        text.insert("1.0", content)
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=(0, 8))

    def save_excel_dialog(self):
        if load_workbook is None:
            messagebox.showerror("Missing dependency", "Install openpyxl first:\n\npip install openpyxl")
            return
        source = self.xlsx_path.get().strip()
        if not source:
            messagebox.showerror("No workbook", "Load the template workbook first.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile="ARIA_AD_User_List_Edited.xlsx",
        )
        if not path:
            return

        try:
            wb = load_workbook(source)
            ws = wb["Users"]
            header_values = [c.value for c in ws[1]]
            colmap = {name: header_values.index(name) + 1 for name in EXPECTED_COLUMNS}

            # Clear old data values while keeping workbook formatting.
            for row in ws.iter_rows(min_row=2, max_row=max(ws.max_row, len(self.rows) + 1)):
                for cell in row:
                    if cell.column <= len(header_values):
                        cell.value = None

            for row_index, data in enumerate(self.rows, start=2):
                for col in EXPECTED_COLUMNS:
                    value = data.get(col, "")
                    if col in BOOL_COLUMNS:
                        value = as_bool(value)
                    ws.cell(row=row_index, column=colmap[col], value=value)

            if "ARIAUsersTable" in ws.tables:
                ws.tables["ARIAUsersTable"].ref = f"A1:Q{max(2, len(self.rows) + 1)}"

            if "Config" in wb.sheetnames:
                cfg_ws = wb["Config"]
                updates = {
                    "DomainFQDN": self.domain_var.get().strip(),
                    "DomainController": self.dc_var.get().strip(),
                    "DefaultOU": self.ou_var.get().strip(),
                    "UPNSuffix": self.upn_var.get().strip(),
                    "StandardGroup": self.standard_group_var.get().strip(),
                    "AdminGroups": self.admin_groups_var.get().strip(),
                }
                for row_index in range(2, cfg_ws.max_row + 1):
                    key = cfg_ws.cell(row_index, 1).value
                    if key in updates:
                        cfg_ws.cell(row_index, 2, updates[key])

            wb.save(path)
            self.write_log(f"Saved edited workbook: {path}")
            messagebox.showinfo("Saved", path)
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))


def main():
    app = App()
    if len(sys.argv) > 1:
        candidate = Path(sys.argv[1])
        if candidate.exists():
            app.after(100, lambda: app.load_excel(str(candidate)))
    app.mainloop()


if __name__ == "__main__":
    main()
