# VMware Workstation Auditor

A graphical PowerShell utility that audits VMware Workstation environments, extracting VM metadata, snapshots, configuration keys, and version details, then displays them in a sortable GUI table.

![image](https://github.com/user-attachments/assets/715afd4a-6863-475a-bf83-afa57a98a9b8)

Affter running:

![image](https://github.com/user-attachments/assets/e7a3337b-da34-4972-a523-ee98a6394f09)

---

## 🔍 Features

- GUI-based directory selection and auditing
- Lists all `.vmx` virtual machine configuration files
- Extracts key-value pairs from `.vmx` files
- Parses snapshot metadata from `.vmsd` files
- Displays VMware version detected on the system
- Sortable table output with detailed VM configuration
- Export audit report to `.CSV`
- Check for updates

---

## 📦 Requirements

- Windows PowerShell 5.1 or later
- VMware Workstation (for actual auditing data)
- Execution policy must permit script execution:
  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
▶️ Usage
Clone or download this repository.

Open PowerShell and run the script:

powershell
Copy
Edit
.\vmware_auditor.ps1
Browse to your Virtual Machines directory or accept the default.

Click Run Audit.

Optionally, Save to CSV for offline reporting.

📁 Output
Each audit entry contains:

Column	Description
VM Name	Name of the virtual machine
Key	Configuration key or snapshot metadata type
Value	Value associated with the key
VM Version	Hardware version (from virtualHW.version)
Snapshot Info	Snapshot UID or metadata key

🛡 License
text
Copy
Edit
Apache License 2.0
Copyright 2025 Richard Wadsworth
See LICENSE for full terms.

🤝 Contributions
Contributions are welcome via pull requests. Bug reports and feature suggestions can be submitted through GitHub Issues.

📞 Contact
For professional inquiries or related scripting needs, please contact:
Richard Wadsworth
IT & Systems Auditor Infomation Assurance Manger 
