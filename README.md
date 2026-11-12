# ARIA v18.x Active Directory User Provisioning GUI

Files:
- `aria_ad_user_gui.py` — Windows/Tkinter GUI.
- `ARIA_AD_User_Template_v18_3.xlsx` — Excel input template with sample standard user and interactive administrator.
- `requirements_aria_ad_gui.txt` — Python dependency (`openpyxl`).
- `run_aria_ad_user_gui.bat` — convenience launcher.

## Manual basis used

The supplied ARIA v18.3 / Eclipse v18.x System Preparation Manual states that ARIA/Eclipse uses Windows Active Directory authentication, account/password management is performed by local IT, and ARIA application rights remain managed in Shared Framework/VSP. It also defines the recommended Varian AD groups and the manual AD user-creation workflow.

The template implements the normal interactive-user workflow:
- Standard ARIA user → `Varian Users`.
- Interactive administrator example → `Varian Users;Varian Administrators`.
- `Varian Users` is expected to be nested into `Varian Application Users` in the recommended group model.

The script deliberately does **not** create or reconfigure the Varian security groups; these should already exist according to the site's validated ARIA security configuration.

## Before use

1. Run on a Windows administration workstation with network access to the AD domain.
2. Install Python 3.
3. Install dependency:
   ```cmd
   py -3 -m pip install -r requirements_aria_ad_gui.txt
   ```
4. Install/enable RSAT **Active Directory Domain Services and Lightweight Directory Services Tools** so that the PowerShell `ActiveDirectory` module is available.
5. Use an account with sufficient AD permissions to create users in the target OU and add users to the required groups.
6. Edit the **Config** sheet in the Excel file for your real domain, OU and group names before execution.

## Recommended workflow

1. Start `run_aria_ad_user_gui.bat` or:
   ```cmd
   py -3 aria_ad_user_gui.py ARIA_AD_User_Template_v18_3.xlsx
   ```
2. Review the loaded rows.
3. Double-click a cell to edit it; add/remove rows as required.
4. Click **Validate**.
5. Click **AD Preflight** to verify domain access, OU, groups and view the domain password/lockout policy.
6. Click **What-If**. This checks that each user does not already exist and that the referenced OU/groups exist, without creating anything.
7. Only then click **EXECUTE CREATE**.
8. After successful AD creation, create/activate/link the user in **Shared Framework / Varian Service Portal** and assign the required ARIA application rights.

## Password modes

- `Prompt` — recommended. Password is entered securely at execution and is not stored in Excel.
- `Generate` — creates a temporary 16-character password; it is shown once after execution and is not written to the audit log.
- `Excel` — reads `InitialPassword` from the workbook. This stores a plaintext password and is discouraged.

For interactive accounts the GUI enforces at least 12 characters and at least 3 character categories. Your local AD policy can be stricter and remains authoritative.

## Safety behavior

- Existing AD accounts are **skipped** and never modified.
- OU and group existence is checked before account creation.
- Passwords are not included in preview or audit log.
- Audit results are appended to `ARIA_AD_User_Creation.log` next to the loaded workbook.
- `PasswordNeverExpires` + `ChangePasswordAtLogon` is rejected as an invalid combination.

## Important

This program was syntax-checked, but it cannot be live-tested against hospital AD. Run **AD Preflight** and **What-If** in your test/T-Box environment first, then validate with your local IT/Varian security procedure before production use.
