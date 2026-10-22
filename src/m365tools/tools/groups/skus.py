"""Readable names for licence SKUs. Graph's ``/subscribedSkus`` gives part numbers like ENTERPRISEPACK; this maps
both the part numbers and the well-known SKU ids to the product names admins use, so licences stay readable
even when the signed-in account may not read ``/subscribedSkus``."""

from __future__ import annotations

# skuId -> (part number, product name). Source: Microsoft's "Product names and service plan identifiers" list.
KNOWN_SKUS: dict[str, tuple[str, str]] = {
    "18181a46-0d4e-45cd-891e-60aabd171b4e": ("STANDARDPACK", "Office 365 E1"),
    "6fd2c87f-b296-42f0-b197-1e91e994b900": ("ENTERPRISEPACK", "Office 365 E3"),
    "c7df2760-2c81-4ef7-b578-5b5392b571df": ("ENTERPRISEPREMIUM", "Office 365 E5"),
    "4b585984-651b-448a-9e53-3b10f069cf7f": ("DESKLESSPACK", "Office 365 F3"),
    "05e9a617-0261-4cee-bb44-138d3ef5d965": ("SPE_E3", "Microsoft 365 E3"),
    "06ebc4ee-1bb5-47dd-8e1b-7f5c2dd87e2e": ("SPE_E5", "Microsoft 365 E5"),
    "66b55226-6b4f-492c-910c-a3b7a3c9d993": ("SPE_F1", "Microsoft 365 F3"),
    "3b555118-da6a-4418-894f-7df1e2096870": ("O365_BUSINESS_ESSENTIALS", "Microsoft 365 Business Basic"),
    "f245ecc8-75af-4f8e-b61f-27d8114de5f3": ("O365_BUSINESS_PREMIUM", "Microsoft 365 Business Standard"),
    "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46": ("SPB", "Microsoft 365 Business Premium"),
    "4b9405b0-7788-4568-add1-99614e613b69": ("EXCHANGESTANDARD", "Exchange Online (Plan 1)"),
    "19ec0d23-8335-4cbd-94ac-6050e30712fa": ("EXCHANGEENTERPRISE", "Exchange Online (Plan 2)"),
    "efccb6f7-5641-4e0e-bd10-b4976e1bf68e": ("EMS", "Enterprise Mobility + Security E3"),
    "26d45bd9-adf1-46cd-a9e1-51e9a5524128": ("EMSPREMIUM", "Enterprise Mobility + Security E5"),
    "4ef96642-f096-40de-a3e9-d83fb2f90211": ("ATP_ENTERPRISE", "Defender for Office 365 (Plan 1)"),
    "e43b5b99-8dfb-405f-9987-dc307f34bcbd": ("MCOEV", "Teams Phone Standard"),
    "0dab259f-bf13-4952-b7f8-7db8f131b28d": ("MCOEV", "Teams Phone Standard"),
    "440eaaa8-b3e0-484b-a8be-62870b9ba70a": ("PHONESYSTEM_VIRTUALUSER", "Teams Phone Resource Account"),
    "6070a4c8-34c6-4937-8dfb-39bbc6397a60": ("MEETING_ROOM", "Teams Rooms Standard"),
    "710779e8-3d4a-4c88-adb9-386c958d1fdf": ("TEAMS_EXPLORATORY", "Teams Exploratory"),
    "c5928f49-12ba-48f7-ada3-0d743a3601d5": ("VISIOCLIENT", "Visio Plan 2"),
    "4b244418-9658-4451-a2b8-b5e2b364e9bd": ("VISIOONLINE_PLAN1", "Visio Plan 1"),
    "53818b1b-4a27-454b-8896-0dba576410e6": ("PROJECTPROFESSIONAL", "Project Plan 3"),
    "09015f9f-377f-4538-bbb5-f75ceb09358a": ("PROJECTPREMIUM", "Project Plan 5"),
    "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235": ("POWER_BI_STANDARD", "Power BI (free)"),
    "f8a1db68-be16-40ed-86d5-cb42ce701560": ("POWER_BI_PRO", "Power BI Pro"),
    "f30db892-07e9-47e9-837c-80727f46fd3d": ("FLOW_FREE", "Power Automate Free"),
    "dcb1a3ae-b33f-4487-846a-a640262fadf4": ("POWERAPPS_VIRAL", "Power Apps Plan 2 Trial"),
    "1f2f344a-700d-42c9-9427-5cea1d5d7ba6": ("STREAM", "Microsoft Stream Trial"),
    "6470687e-a428-4b7a-bef2-8a291ad947c9": ("WINDOWS_STORE", "Windows Store for Business"),
    "7cfd9a2b-e110-4c39-bf20-c6a3f36a3121": ("WIN10_PRO_ENT_SUB", "Windows 10/11 Enterprise E3"),
    "8c4ce438-32a7-4ac5-91a6-e22ae08d9c8b": ("RIGHTSMANAGEMENT_ADHOC", "Rights Management Adhoc"),
    "90d8b3f8-712e-4f7b-aa1e-62e7ae6cbe96": ("SMB_APPS", "Business Apps (free)"),
    "639dec6b-bb19-468b-871c-c5c441c4b0cb": ("Microsoft_365_Copilot", "Microsoft 365 Copilot"),
    "3f9f06f5-3c31-472c-985f-62d9c10ec167": ("Power_Pages_vTrial_for_Makers", "Power Pages trial"),
    "e8f81a67-bd96-4074-b108-cf193eb9433b": ("WIN_DEF_ATP", "Defender for Endpoint"),
    "16a55f2f-ff35-4cd5-9146-fb784e3761a5": ("MDE_SMB", "Defender for Business"),
    "61e6bd70-fbdb-4deb-82ea-912842f39431": ("DYN365_CS_CHAT", "Dynamics 365 Customer Service Chat"),
}

_PART_TO_NAME: dict[str, str] = {part: name for part, name in KNOWN_SKUS.values()}


def licence_name(sku_id: str, part_numbers: dict[str, str] | None = None) -> str:
    """Readable product name for ``sku_id``: the known table first, then the tenant's part number, then a short id."""
    known = KNOWN_SKUS.get(sku_id.lower())
    if known:
        return known[1]
    part = (part_numbers or {}).get(sku_id)
    if part:
        return _PART_TO_NAME.get(part, part)
    return f"licence {sku_id[:8]}"
