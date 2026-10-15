# azure-tf-asset-audit

Repeatable, read-only comparison of an Azure subscription against Terraform state, producing an Excel asset registry.

## Setup (once)

**Prerequisites:** Python 3.9+, [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli), [Terraform](https://developer.hashicorp.com/terraform/install).

### 1. Create and activate a virtual environment

Windows (PowerShell):
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
# If scripts are blocked:  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Windows (cmd):
```bat
python -m venv venv
venv\Scripts\activate.bat
```

Linux / macOS / WSL:
```bash
python3 -m venv venv
source venv/bin/activate
```

> A venv is not portable between operating systems. If you use both Windows and WSL on the same folder, create a separate one for each (e.g. `venv-wsl`; the `.gitignore` covers `venv/` and `.venv/`, so add any other name you choose).

Leave the venv with `deactivate`.

### 2. Install dependencies and log in
```
pip install -r requirements.txt
az login
az extension add --name resource-graph   # optional but faster/more complete
terraform init                            # in each TF root, so state is readable
```

## Run
```powershell
python azure_tf_asset_audit.py --subscription <SUB_ID> --tf-dir C:\path\to\terraform
# Linux/WSL path example: --tf-dir /mnt/c/path/to/terraform
# several TF roots:  --tf-dir envs\prod --tf-dir envs\shared
# offline / CI:      --inventory-json output\inventory_X.json --tf-json output\tfshow_X.json
```
Output in `./output` (git-ignored, as it contains real tenant data): `azure_asset_registry_<timestamp>.xlsx` plus raw JSON snapshots (keep these for history/diffing).

## Workbook sheets
- **Summary** – counts and Terraform coverage %.
- **Registry** – every Azure resource: status, type, RG, location, SKU, tags, created, TF address, ID.
- **Unmanaged** – in Azure, not in Terraform (candidates to import or delete).
- **By type / By resource group** – roll-ups.
- **In TF not in Azure** – state entries with no Azure resource (deleted outside TF, stale state).
- **TF non-Azure** – azuread/random/other-subscription resources.

## Matching rules
Lower-cased Azure resource ID. A TF child object (subnet, diagnostic setting, …) marks its parent as Managed. Platform noise types are `Ignored` (see `DEFAULT_IGNORE_TYPES`; add `--ignore-type`).

## Limits
- Compares against **state**, not `.tf` code. For config drift (attribute differences) also run `terraform plan -detailed-exitcode` per root (exit 2 = drift/changes).
- Resource Graph `resources` table excludes role assignments, policy, Entra objects and some child types.
- Needs Reader on the subscription plus access to the TF backend.

## License

[MIT](LICENSE)
