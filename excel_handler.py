from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
import pandas as pd


class ExcelHandler:

    def __init__(self):
        self.workbook = None
        self.sheet = None

    # -------------------------------------------------------
    # Load Workbook
    # -------------------------------------------------------

    def load(self, path):
        self.workbook = load_workbook(path)
        self.sheet = self.workbook.active

        return self.workbook, self.sheet

    # -------------------------------------------------------
    # Ensure Customers Sheet Exists
    # -------------------------------------------------------

    def ensure_customers_sheet(self):

        if "Customers" in self.workbook.sheetnames:

            customers_sheet = self.workbook[
                "Customers"
            ]

            return customers_sheet


        # Create Customers sheet
        customers_sheet = self.workbook.create_sheet(
            "Customers"
        )


        # Headers
        customers_sheet["A1"] = "Customer Code"
        customers_sheet["B1"] = "Customer Name"


        # Header formatting
        header_fill = PatternFill(
            fill_type="solid",
            fgColor="1F2937"
        )

        header_font = Font(
            color="FFFFFF",
            bold=True
        )

        header_alignment = Alignment(
            horizontal="center",
            vertical="center"
        )


        for cell in customers_sheet[1]:

            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment


        customers_sheet.column_dimensions[
            "A"
        ].width = 20

        customers_sheet.column_dimensions[
            "B"
        ].width = 35


        # Copy existing customers from current sheet
        for row in range(
            2,
            self.sheet.max_row + 1
        ):

            code = self.sheet.cell(
                row,
                1
            ).value

            name = self.sheet.cell(
                row,
                2
            ).value

            if code is None:
                continue

            customers_sheet.append([
                str(code).strip(),
                (
                    str(name).strip()
                    if name is not None
                    else ""
                )
            ])


        return customers_sheet


    # -------------------------------------------------------
    # Get Customers From Customers Sheet
    # -------------------------------------------------------

    def get_active_customers(self):

        customers_sheet = (
            self.ensure_customers_sheet()
        )

        customers = []


        for row in range(
            2,
            customers_sheet.max_row + 1
        ):

            code = customers_sheet.cell(
                row,
                1
            ).value

            name = customers_sheet.cell(
                row,
                2
            ).value


            if code is None:
                continue


            code = str(code).strip()


            if not code:
                continue


            customers.append({
                "Customer Code": code,
                "Customer Name": (
                    str(name).strip()
                    if name is not None
                    else ""
                )
            })


        customer_df = pd.DataFrame(
            customers,
            columns=[
                "Customer Code",
                "Customer Name"
            ]
        )


        customer_dict = dict(
            zip(
                customer_df[
                    "Customer Code"
                ],
                customer_df[
                    "Customer Name"
                ]
            )
        )


        return (
            customer_df,
            customer_dict
        )

    # -------------------------------------------------------
    # Add / Update Customer
    # -------------------------------------------------------

    def add_or_update_customer(
        self,
        customer_code,
        customer_name
    ):

        customers_sheet = (
            self.ensure_customers_sheet()
        )


        customer_code = (
            str(customer_code)
            .strip()
            .upper()
        )

        customer_name = (
            str(customer_name)
            .strip()
        )


        if not customer_code:
            raise ValueError(
                "Customer code is required."
            )


        if not customer_name:
            raise ValueError(
                "Customer name is required."
            )


        # ---------------------------------------------------
        # Check if customer already exists
        # ---------------------------------------------------

        existing_row = None
        old_name = None


        for row in range(
            2,
            customers_sheet.max_row + 1
        ):

            existing_code = (
                customers_sheet.cell(
                    row,
                    1
                ).value
            )


            if (
                existing_code is not None
                and str(existing_code).strip().upper()
                == customer_code
            ):

                existing_row = row

                old_name = (
                    customers_sheet.cell(
                        row,
                        2
                    ).value
                )

                break


        # ---------------------------------------------------
        # UPDATE existing customer
        # ---------------------------------------------------

        if existing_row is not None:

            customers_sheet.cell(
                existing_row,
                2
            ).value = customer_name


            # Update current worksheet name
            current_row = (
                self.find_customer_row(
                    customer_code
                )
            )


            if current_row is not None:

                self.sheet.cell(
                    current_row,
                    2
                ).value = customer_name


            return {
                "action": "Updated",
                "old_name": (
                    str(old_name).strip()
                    if old_name is not None
                    else None
                ),
                "new_name": customer_name
            }


        # ---------------------------------------------------
        # ADD new customer
        # ---------------------------------------------------

        new_customer_row = (
            customers_sheet.max_row + 1
        )


        customers_sheet.cell(
            new_customer_row,
            1
        ).value = customer_code

        customers_sheet.cell(
            new_customer_row,
            2
        ).value = customer_name


        return {
            "action": "Added",
            "old_name": None,
            "new_name": customer_name
        }

    # -------------------------------------------------------
    # Delete Customer
    # -------------------------------------------------------

    def delete_customer(
        self,
        customer_code
    ):

        customers_sheet = (
            self.ensure_customers_sheet()
        )

        customer_code = (
            str(customer_code)
            .strip()
            .upper()
        )

        existing_row = None
        old_name = None

        # ---------------------------------------------------
        # Find customer in Customers master
        # ---------------------------------------------------

        for row in range(
            2,
            customers_sheet.max_row + 1
        ):

            existing_code = (
                customers_sheet.cell(
                    row,
                    1
                ).value
            )

            if (
                existing_code is not None
                and str(existing_code).strip().upper()
                == customer_code
            ):

                existing_row = row

                old_name = (
                    customers_sheet.cell(
                        row,
                        2
                    ).value
                )

                break

        if existing_row is None:

            raise ValueError(
                "Customer not found."
            )

        # ---------------------------------------------------
        # Remove from Customers master
        # ---------------------------------------------------

        customers_sheet.delete_rows(
            existing_row,
            1
        )

        # ---------------------------------------------------
        # Current worksheet
        #
        # KEEP:
        #   Customer Code
        #
        # CLEAR:
        #   Customer Name
        #   All collection amounts
        # ---------------------------------------------------

        current_row = (
            self.find_customer_row(
                customer_code
            )
        )

        if current_row is not None:

            # Keep Customer Code
            self.sheet.cell(
                current_row,
                1
            ).value = customer_code

            # Remove Customer Name
            self.sheet.cell(
                current_row,
                2
            ).value = None

            # Remove all current-month amounts
            for column in range(
                3,
                self.sheet.max_column + 1
            ):

                self.sheet.cell(
                    current_row,
                    column
                ).value = None

        # ---------------------------------------------------
        # Return information for audit
        # ---------------------------------------------------

        return {
            "action": "Deleted",

            "old_name": (
                str(old_name).strip()
                if old_name is not None
                else None
            ),

            "new_name": None
        }

    # -------------------------------------------------------
    # Convert Excel Sheet to DataFrame
    # -------------------------------------------------------

    def dataframe(self):

        data = self.sheet.values

        cols = next(data)

        df = pd.DataFrame(
            data,
            columns=cols
        )

        # Columns from the 3rd column onward are
        # daily collection amounts.
        amt = df.columns[2:]

        df[amt] = (
            df[amt]
            .apply(
                pd.to_numeric,
                errors="coerce"
            )
            .round(2)
        )

        return (
            df.astype(object)
            .where(
                pd.notna(df),
                "-"
            )
        )

    # -------------------------------------------------------
    # Find Customer Row
    # -------------------------------------------------------

    def find_customer_row(self, code):

        for r in range(
            2,
            self.sheet.max_row + 1
        ):

            cell_value = self.sheet.cell(
                r,
                1
            ).value

            if (
                str(cell_value).strip()
                == str(code).strip()
            ):
                return r

        return None

    # -------------------------------------------------------
    # Find Day Column
    # -------------------------------------------------------

    def find_date_column(self, day):

        for c in range(
            3,
            self.sheet.max_column + 1
        ):

            cell_value = self.sheet.cell(
                1,
                c
            ).value

            if cell_value == day:
                return c

        return None

    # -------------------------------------------------------
    # Update Collection Amount
    # -------------------------------------------------------

    def update(
        self,
        code,
        day,
        amount
    ):

        row = self.find_customer_row(
            code
        )

        column = self.find_date_column(
            day
        )

        if row is None:
            raise ValueError(
                "Customer not found"
            )

        if column is None:
            raise ValueError(
                "Date not found"
            )

        self.sheet.cell(
            row,
            column
        ).value = amount

        return row, column

    # -------------------------------------------------------
    # Save Workbook
    # -------------------------------------------------------

    def save(self, path):

        self.workbook.save(path)

    # -------------------------------------------------------
    # Close Workbook
    # -------------------------------------------------------

    def close(self):

        if self.workbook:
            self.workbook.close()

    # -------------------------------------------------------
    # Get Customer Data
    # -------------------------------------------------------

    def get_customer_data(self):
        """
        Returns:

        customer_df
            DataFrame containing:
            Customer Code
            Customer Name

        customer_dict
            Dictionary:
            Customer Code -> Customer Name
        """

        df = self.dataframe()

        customer_df = df.iloc[:, 0:2].copy()

        customer_df.columns = [
            "Customer Code",
            "Customer Name"
        ]

        customer_df = customer_df.dropna(
            subset=[
                "Customer Code"
            ]
        )

        customer_df[
            "Customer Code"
        ] = (
            customer_df[
                "Customer Code"
            ].astype(str)
        )

        customer_dict = dict(
            zip(
                customer_df[
                    "Customer Code"
                ],
                customer_df[
                    "Customer Name"
                ]
            )
        )

        return (
            customer_df,
            customer_dict
        )