import os
import tempfile

from excel_handler import ExcelHandler
from workbook_storage import (
    download_master_template,
    upload_master_template
)


def _customer_map(customer_df):
    required = {"Customer Code", "Customer Name"}
    if not required.issubset(customer_df.columns):
        raise ValueError("Customer Code/Customer Name columns are missing")

    customers = {}
    for _, row in customer_df.iterrows():
        raw_code = str(row["Customer Code"]).strip()
        name = str(row["Customer Name"]).strip()
        if not raw_code or raw_code.lower() == "nan":
            continue
        code = raw_code.upper()
        if code in customers and customers[code] != name:
            raise ValueError(f"Duplicate customer code: {code}")
        customers[code] = name
    return customers


def sync_master_template_from_current(current_handler, bundled_template_path):
    """Reconcile ONLY customer records from the current sheet into the master.

    This deliberately never copies the current month's workbook into master;
    therefore tea amounts for a month are not used to make a new template.
    """
    current_df, _ = current_handler.get_active_customers()
    current_customers = _customer_map(current_df)

    # Prevent accidental removal of every template customer if the source
    # cannot be read or is unexpectedly empty.
    if not current_customers:
        raise ValueError("Current customer list is empty. Template was not changed.")

    master_bytes = download_master_template()
    master_is_new = master_bytes is None

    if master_is_new:
        with open(bundled_template_path, "rb") as template_file:
            master_bytes = template_file.read()

    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            delete=False, suffix=".xlsx"
        ) as temp:
            temp.write(master_bytes)
            temp_path = temp.name

        template_handler = ExcelHandler()
        template_handler.load(temp_path)
        template_df, _ = template_handler.get_active_customers()
        existing_customers = _customer_map(template_df)

        removed = 0
        added = 0
        renamed = 0

        # Delete customers no longer present in the CURRENT sheet.
        # Older months are never opened or changed.
        for code in existing_customers:
            if code not in current_customers:
                template_handler.delete_customer(code)
                removed += 1

        # Add new customers and update changed names.
        for code, name in current_customers.items():
            if code not in existing_customers:
                template_handler.add_or_update_customer(code, name)
                added += 1
            elif existing_customers[code] != name:
                template_handler.add_or_update_customer(code, name)
                renamed += 1

        if master_is_new or added or removed or renamed:
            template_handler.save(temp_path)
            with open(temp_path, "rb") as template_file:
                upload_master_template(template_file.read())

        return {"added": added, "updated": renamed, "deleted": removed}

    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)
