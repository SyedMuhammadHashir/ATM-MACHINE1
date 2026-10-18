
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.image import Image as XLImage
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt, RGBColor
import os
import json
import re
from datetime import datetime


class ATM:

    # =========================================================
    # COLORS
    # =========================================================

    RESET = "\033[0m"
    BOLD = "\033[1m"
    RED = "\033[91m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    BLUE = "\033[94m"
    MAGENTA = "\033[95m"
    CYAN = "\033[96m"
    WHITE = "\033[97m"

    def color(self, text, color):
        return f"{color}{text}{self.RESET}"

    def show_box(self, title, lines, color=None):
        color = color or self.CYAN
        width = 58

        print()
        print(self.color("╔" + "═" * width + "╗", color))
        print(self.color("║" + title.center(width) + "║", color))
        print(self.color("╠" + "═" * width + "╣", color))

        for line in lines:
            line = str(line)

            if len(line) > width - 2:
                line = line[:width - 2]

            print(
                self.color(
                    "║ " + line.ljust(width - 2) + " ║",
                    color
                )
            )

        print(self.color("╚" + "═" * width + "╝", color))

    # =========================================================
    # INITIALIZATION
    # =========================================================

    def __init__(self):

        self.accounts_folder = "ATM_Accounts"

        os.makedirs(
            self.accounts_folder,
            exist_ok=True
        )

        self.registry_file = os.path.join(
            self.accounts_folder,
            "account_registry.json"
        )

        self.registry = self.load_registry()

        self.name = None
        self.pin = None
        self.account_id = None
        self.balance = 0
        self.transactions = []
        self.account = {}
        self.account_info = ("Saving Account", None)

        self.excel_file = None
        self.word_file = None

    # =========================================================
    # REGISTRY
    # =========================================================

    def load_registry(self):

        if not os.path.exists(self.registry_file):
            return {
                "accounts": {},
                "used_pins": []
            }

        try:

            with open(
                self.registry_file,
                "r",
                encoding="utf-8"
            ) as file:

                data = json.load(file)

            if isinstance(
                data.get("accounts"),
                dict
            ):

                data.setdefault(
                    "used_pins",
                    []
                )

                return data

            old_accounts = data.get(
                "accounts",
                []
            )

            new_accounts = {}

            for account in old_accounts:

                name = str(
                    account.get(
                        "name",
                        ""
                    )
                ).strip()

                if name:

                    new_accounts[name.lower()] = {
                        "Name": name,
                        "PIN": str(
                            account.get(
                                "pin",
                                ""
                            )
                        ),
                        "Balance": 10000,
                        "Account ID": account.get(
                            "account_id",
                            ""
                        ),
                        "created_on": account.get(
                            "created_on",
                            datetime.now().strftime(
                                "%Y-%m-%d %H:%M:%S"
                            )
                        )
                    }

            return {
                "accounts": new_accounts,
                "used_pins": data.get(
                    "used_pins",
                    []
                )
            }

        except Exception as e:

            print(
                f"Error loading account registry: {e}"
            )

            return {
                "accounts": {},
                "used_pins": []
            }

    def save_registry(self):

        try:

            with open(
                self.registry_file,
                "w",
                encoding="utf-8"
            ) as file:

                json.dump(
                    self.registry,
                    file,
                    indent=4
                )

        except Exception as e:

            print(
                f"Error saving account registry: {e}"
            )

    def generate_account_id(self):

        numbers = []

        for account in self.registry["accounts"].values():

            account_id = str(
                account.get(
                    "Account ID",
                    ""
                )
            )

            match = re.search(
                r"ATM-(\d+)",
                account_id
            )

            if match:

                numbers.append(
                    int(match.group(1))
                )

        next_number = max(
            numbers,
            default=0
        ) + 1

        return f"ATM-{next_number:03d}"

    def safe_filename(self, name):

        safe_name = re.sub(
            r'[<>:"/\\|?*]',
            "",
            name
        )

        safe_name = safe_name.replace(
            " ",
            "_"
        )

        return (
            safe_name
            if safe_name
            else "Account"
        )

    # =========================================================
    # CREATE ACCOUNT
    # =========================================================

    def create_account(self):

        print("\n" + "=" * 55)
        print("              CREATE NEW ATM ACCOUNT")
        print("=" * 55)

        name = input(
            "Enter Account Holder Name: "
        ).strip()

        while not name:

            print(
                "Name cannot be empty!"
            )

            name = input(
                "Enter Account Holder Name: "
            ).strip()

        username = name.lower()

        if username in self.registry["accounts"]:

            print(
                "\nAccount already exists with this name!"
            )

            print(
                "Please use Login if this is your account."
            )

            return False

        pin = input(
            "Enter Your PIN (4 digits): "
        ).strip()

        while (
            len(pin) != 4
            or not pin.isdigit()
        ):

            print(
                "PIN must be exactly 4 digits!"
            )

            pin = input(
                "Enter Your PIN (4 digits): "
            ).strip()

        if pin in self.registry["used_pins"]:

            print(
                "\nThis PIN is already being used "
                "by another account."
            )

            print(
                "Please choose a different PIN."
            )

            return False

        account_id = self.generate_account_id()

        balance = 10000

        self.registry["accounts"][username] = {

            "Name": name,

            "PIN": pin,

            "Balance": balance,

            "Account ID": account_id,

            "Account Type": "Saving Account",

            "created_on":
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
        }

        self.registry["used_pins"].append(
            pin
        )

        self.save_registry()

        self.load_account(username)

        self.transactions = []

        self.initialize_files()

        print(
            self.color(
                "\n╔" + "═" * 55 + "╗",
                self.GREEN
            )
        )

        print(
            self.color(
                "║" +
                "ACCOUNT CREATED SUCCESSFULLY".center(55) +
                "║",
                self.GREEN
            )
        )

        print(
            self.color(
                "╚" + "═" * 55 + "╝",
                self.GREEN
            )
        )

        print(
            f"Account Holder : {self.name}"
        )

        print(
            f"Account ID     : {self.account_id}"
        )

        print(
            f"Account Type   : {self.account_info[0]}"
        )

        print(
            f"Initial Balance: Rs.{self.balance:,.2f}"
        )

        print("=" * 55)

        return True

    # =========================================================
    # LOGIN
    # =========================================================

    def login(self):

        print("\n" + "=" * 55)
        print("                    ATM LOGIN")
        print("=" * 55)

        name = input(
            "Enter Account Name: "
        ).strip()

        pin = input(
            "Enter PIN: "
        ).strip()

        username = name.lower()

        if username not in self.registry["accounts"]:

            print(
                "\nAccount not found!"
            )

            return False

        account = self.registry["accounts"][username]

        if str(account["PIN"]) != pin:

            print(
                "\nWrong PIN!"
            )

            return False

        self.load_account(username)

        print("\n" + "=" * 55)
        print("                 LOGIN SUCCESSFUL")
        print("=" * 55)

        print(
            f"Welcome, {self.name}!"
        )

        print(
            f"Account ID: {self.account_id}"
        )

        print(
            f"Balance: Rs.{self.balance:,.2f}"
        )

        print("=" * 55)

        return True

    # =========================================================
    # LOAD ACCOUNT
    # =========================================================

    def load_account(self, username):

        account = self.registry["accounts"][username]

        self.name = account["Name"]

        self.pin = str(
            account["PIN"]
        )

        self.balance = float(
            account.get(
                "Balance",
                10000
            )
        )

        self.account_id = account[
            "Account ID"
        ]

        self.account = account

        self.account_info = (
            account.get(
                "Account Type",
                "Saving Account"
            ),
            self.account_id
        )

        safe_name = self.safe_filename(
            self.name
        )

        self.excel_file = os.path.join(
            self.accounts_folder,
            f"{self.account_id}_{safe_name}_transactions.xlsx"
        )

        self.word_file = os.path.join(
            self.accounts_folder,
            f"{self.account_id}_{safe_name}_report.docx"
        )

        self.load_transactions()

    # =========================================================
    # SAVE CURRENT ACCOUNT
    # =========================================================

    def save_current_account(self):

        username = self.name.lower()

        if username in self.registry["accounts"]:

            self.registry["accounts"][username][
                "Balance"
            ] = self.balance

            self.registry["accounts"][username][
                "PIN"
            ] = self.pin

            self.save_registry()

    # =========================================================
    # LOAD TRANSACTIONS
    # =========================================================

    def load_transactions(self):

        self.transactions = []

        if not os.path.exists(
            self.excel_file
        ):
            return

        try:

            wb = load_workbook(
                self.excel_file,
                data_only=True
            )

            ws = wb.active

            for row in ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not row:
                    continue

                transaction = row[1]

                amount = row[2]

                balance = row[3]

                if transaction in (
                    "Deposit",
                    "Withdraw"
                ):

                    self.transactions.append(
                        f"{transaction} : "
                        f"Rs.{float(amount):,.2f} | "
                        f"Balance: "
                        f"Rs.{float(balance):,.2f}"
                    )

            wb.close()

        except Exception as e:

            print(
                f"Could not load previous transactions: {e}"
            )

    # =========================================================
    # INITIALIZE FILES
    # =========================================================

    def initialize_files(self):

        if not os.path.exists(
            self.excel_file
        ):

            self.create_excel()

        if not os.path.exists(
            self.word_file
        ):

            self.create_word()

    # =========================================================
    # CREATE EXCEL
    # =========================================================

    def create_excel(self):

        try:

            wb = Workbook()

            ws = wb.active

            ws.title = "ATM Transactions"

            headers = [
                "Date",
                "Transaction",
                "Amount",
                "Balance"
            ]

            ws.append(headers)

            header_font = Font(
                name="Arial",
                size=12,
                bold=True,
                color="FFFFFF"
            )

            header_fill = PatternFill(
                start_color="1F4E78",
                end_color="1F4E78",
                fill_type="solid"
            )

            thin_border = Border(
                left=Side(
                    style="thin",
                    color="000000"
                ),
                right=Side(
                    style="thin",
                    color="000000"
                ),
                top=Side(
                    style="thin",
                    color="000000"
                ),
                bottom=Side(
                    style="thin",
                    color="000000"
                )
            )

            for col in range(1, 5):

                cell = ws.cell(
                    row=1,
                    column=col
                )

                cell.font = header_font

                cell.fill = header_fill

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

                cell.border = thin_border

            ws.column_dimensions["A"].width = 22
            ws.column_dimensions["B"].width = 18
            ws.column_dimensions["C"].width = 18
            ws.column_dimensions["D"].width = 18

            ws.append([
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
                "Initial Balance",
                10000,
                self.balance
            ])

            for col in range(1, 5):

                cell = ws.cell(
                    row=2,
                    column=col
                )

                cell.border = thin_border

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            wb.save(
                self.excel_file
            )

            wb.close()

        except Exception as e:

            print(
                f"Error creating Excel: {e}"
            )

    # =========================================================
    # CREATE WORD
    # =========================================================

    def create_word(self):

        try:

            doc = Document()

            heading = doc.add_heading(
                "ATM Transaction Report",
                0
            )

            heading.alignment = 1

            doc.add_paragraph(
                f"Account Holder: {self.name}"
            )

            doc.add_paragraph(
                f"Account Type: {self.account_info[0]}"
            )

            doc.add_paragraph(
                f"Account ID: {self.account_id}"
            )

            doc.add_paragraph(
                "Created on: " +
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            )

            doc.add_paragraph(
                "-" * 60
            )

            doc.add_paragraph(
                f"Initial Balance: "
                f"Rs.{self.balance:,.2f}"
            )

            doc.add_paragraph(
                "Transaction History"
            )

            doc.add_paragraph(
                "-" * 60
            )

            doc.save(
                self.word_file
            )

        except Exception as e:

            print(
                f"Error creating Word: {e}"
            )

    # =========================================================
    # SAVE EXCEL TRANSACTION
    # =========================================================

    def save_excel(
        self,
        transaction,
        amount
    ):

        try:

            if not os.path.exists(
                self.excel_file
            ):

                self.create_excel()

            wb = load_workbook(
                self.excel_file
            )

            ws = wb.active

            ws.append([
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
                transaction,
                amount,
                self.balance
            ])

            thin_border = Border(
                left=Side(
                    style="thin",
                    color="000000"
                ),
                right=Side(
                    style="thin",
                    color="000000"
                ),
                top=Side(
                    style="thin",
                    color="000000"
                ),
                bottom=Side(
                    style="thin",
                    color="000000"
                )
            )

            if transaction == "Deposit":

                fill = PatternFill(
                    start_color="C6EFCE",
                    end_color="C6EFCE",
                    fill_type="solid"
                )

                font_color = "006100"

            else:

                fill = PatternFill(
                    start_color="FFC7CE",
                    end_color="FFC7CE",
                    fill_type="solid"
                )

                font_color = "9C0006"

            last_row = ws.max_row

            for col in range(1, 5):

                cell = ws.cell(
                    row=last_row,
                    column=col
                )

                cell.border = thin_border

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

                cell.font = Font(
                    name="Arial",
                    size=10,
                    color=font_color
                )

                cell.fill = fill

            wb.save(
                self.excel_file
            )

            wb.close()

            print(
                f"{transaction} saved to Excel!"
            )

        except PermissionError:

            print(
                "Cannot save Excel. "
                "Please close the Excel file first."
            )

        except Exception as e:

            print(
                f"Error saving Excel: {e}"
            )

    # =========================================================
    # SAVE WORD TRANSACTION
    # =========================================================

    def save_word(
        self,
        transaction,
        amount
    ):

        try:

            if not os.path.exists(
                self.word_file
            ):

                self.create_word()

            doc = Document(
                self.word_file
            )

            p = doc.add_paragraph()

            run1 = p.add_run(
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            )

            run1.bold = True

            p.add_run(" | ")

            run2 = p.add_run(
                transaction
            )

            run2.bold = True

            if transaction == "Deposit":

                run2.font.color.rgb = RGBColor(
                    0,
                    128,
                    0
                )

            else:

                run2.font.color.rgb = RGBColor(
                    255,
                    0,
                    0
                )

            p.add_run(
                f" | Amount = Rs.{amount:,.2f}"
            )

            p.add_run(
                f" | Balance = Rs.{self.balance:,.2f}"
            )

            doc.save(
                self.word_file
            )

            print(
                f"{transaction} saved to Word!"
            )

        except PermissionError:

            print(
                "Cannot save Word. "
                "Please close the Word file first."
            )

        except Exception as e:

            print(
                f"Error saving Word: {e}"
            )

    # =========================================================
    # DEPOSIT
    # =========================================================

    def deposit(self):

        try:

            amount = int(
                input(
                    "Enter Deposit Amount: "
                )
            )

            if amount <= 0:

                print(
                    "Invalid amount!"
                )

                return

            old_balance = self.balance

            self.balance += amount

            self.account["Balance"] = (
                self.balance
            )

            self.save_current_account()

            self.transactions.append(
                f"Deposit : "
                f"Rs.{amount:,.2f} | "
                f"Balance: "
                f"Rs.{self.balance:,.2f}"
            )

            print(
                f"\nOld Balance: "
                f"Rs.{old_balance:,.2f}"
            )

            print(
                f"New Balance: "
                f"Rs.{self.balance:,.2f}"
            )

            self.save_excel(
                "Deposit",
                amount
            )

            self.save_word(
                "Deposit",
                amount
            )

            print(
                self.color(
                    "\n✅ Deposit Successful!",
                    self.GREEN
                )
            )

        except ValueError:

            print(
                "Please enter numbers only."
            )

    # =========================================================
    # WITHDRAW
    # =========================================================

    def withdraw(self):

        try:

            amount = int(
                input(
                    "Enter Withdraw Amount: "
                )
            )

            if amount <= 0:

                print(
                    "Invalid amount!"
                )

                return

            if amount > self.balance:

                print(
                    f"Insufficient Balance! "
                    f"Available: "
                    f"Rs.{self.balance:,.2f}"
                )

                return

            old_balance = self.balance

            self.balance -= amount

            self.account["Balance"] = (
                self.balance
            )

            self.save_current_account()

            self.transactions.append(
                f"Withdraw : "
                f"Rs.{amount:,.2f} | "
                f"Balance: "
                f"Rs.{self.balance:,.2f}"
            )

            print(
                f"\nOld Balance: "
                f"Rs.{old_balance:,.2f}"
            )

            print(
                f"New Balance: "
                f"Rs.{self.balance:,.2f}"
            )

            self.save_excel(
                "Withdraw",
                amount
            )

            self.save_word(
                "Withdraw",
                amount
            )

            print(
                self.color(
                    "\n✅ Withdrawal Successful!",
                    self.GREEN
                )
            )

        except ValueError:

            print(
                "Enter numbers only."
            )

    # =========================================================
    # CHECK BALANCE
    # =========================================================

    def check_balance(self):

        self.show_box(
            "💰 ACCOUNT BALANCE",
            [
                f"👤 Account Holder : {self.name}",
                f"🆔 Account ID     : {self.account_id}",
                f"🏦 Account Type   : {self.account_info[0]}",
                f"💵 Current Balance: Rs.{self.balance:,.2f}"
            ],
            self.GREEN
        )

        print(
            self.color(
                "\n1. 📊 Export Balance to Excel",
                self.CYAN
            )
        )

        print(
            self.color(
                "2. 📄 Export Balance to Word",
                self.YELLOW
            )
        )

        print(
            self.color(
                "3. ↩ Back",
                self.WHITE
            )
        )

        choice = input(
            self.color(
                "👉 Enter choice: ",
                self.CYAN
            )
        ).strip()

        if choice == "1":

            self.export_balance_excel()

        elif choice == "2":

            self.export_balance_word()

    # =========================================================
    # GET TOTALS
    # =========================================================

    def get_transaction_totals(self):

        total_deposit = 0.0

        total_withdraw = 0.0

        for transaction in self.transactions:

            try:

                parts = transaction.split(
                    " | "
                )

                first_part = parts[0]

                transaction_type, amount_text = (
                    first_part.split(
                        " : ",
                        1
                    )
                )

                amount = float(
                    amount_text
                    .replace("Rs.", "")
                    .replace(",", "")
                    .strip()
                )

                if transaction_type == "Deposit":

                    total_deposit += amount

                elif transaction_type == "Withdraw":

                    total_withdraw += amount

            except (
                ValueError,
                IndexError
            ):

                continue

        return (
            total_deposit,
            total_withdraw
        )

    # =========================================================
    # HISTORY
    # =========================================================

    def history(self):

        lines = [
            f"👤 Account Holder : {self.name}",
            f"🆔 Account ID     : {self.account_id}",
            f"💰 Initial Balance: Rs.10,000.00",
            f"💵 Current Balance: Rs.{self.balance:,.2f}"
        ]

        self.show_box(
            "📜 TRANSACTION HISTORY",
            lines,
            self.MAGENTA
        )

        # -----------------------------------------------------
        # TRANSACTION HISTORY FIRST
        # -----------------------------------------------------

        if not self.transactions:

            print(
                self.color(
                    "\nNo transactions found.",
                    self.YELLOW
                )
            )

        else:

            print(
                self.color(
                    "\n┌────┬────────────┬──────────────────────────────┐",
                    self.CYAN
                )
            )

            print(
                self.color(
                    "│ No │ Type       │ Amount / Balance            │",
                    self.CYAN
                )
            )

            print(
                self.color(
                    "├────┼────────────┼──────────────────────────────┤",
                    self.CYAN
                )
            )

            for i, transaction in enumerate(
                self.transactions,
                1
            ):

                parts = transaction.split(
                    " | "
                )

                first = parts[0]

                transaction_type, amount = (
                    first.split(
                        " : ",
                        1
                    )
                )

                balance = parts[1].replace(
                    "Balance: ",
                    ""
                )

                row = (
                    f"│ {i:2d} "
                    f"│ {transaction_type:<10} "
                    f"│ {amount:<12} / "
                    f"{balance:<12} │"
                )

                row_color = (
                    self.GREEN
                    if transaction_type == "Deposit"
                    else self.RED
                )

                print(
                    self.color(
                        row,
                        row_color
                    )
                )

            print(
                self.color(
                    "└────┴────────────┴──────────────────────────────┘",
                    self.CYAN
                )
            )

            # -------------------------------------------------
            # ONE SUMMARY CHART BELOW HISTORY
            # -------------------------------------------------

            self.show_transaction_chart()

        print()

        print(
            self.color(
                "1. 📊 Export History to Excel",
                self.CYAN
            )
        )

        print(
            self.color(
                "2. 📄 Export History to Word",
                self.YELLOW
            )
        )

        print(
            self.color(
                "3. ↩ Back",
                self.WHITE
            )
        )

        choice = input(
            self.color(
                "👉 Enter choice: ",
                self.CYAN
            )
        ).strip()

        if choice == "1":

            self.export_history_excel()

        elif choice == "2":

            self.export_history_word()

    # =========================================================
    # ONE CHART BELOW HISTORY
    # =========================================================

    def show_transaction_chart(self):

        print()

        print(
            self.color(
                "📊 DEPOSIT vs WITHDRAW SUMMARY",
                self.BLUE
            )
        )

        print(
            self.color(
                "─" * 58,
                self.BLUE
            )
        )

        total_deposit, total_withdraw = (
            self.get_transaction_totals()
        )

        total = (
            total_deposit +
            total_withdraw
        )

        if total == 0:

            print(
                self.color(
                    "No chart data available.",
                    self.YELLOW
                )
            )

            return

        # -----------------------------------------------------
        # ONE SIMPLE PIE-STYLE CONSOLE SUMMARY
        # -----------------------------------------------------

        print()

        print(
            self.color(
                f"🟢 Deposit  : Rs.{total_deposit:,.2f}",
                self.GREEN
            )
        )

        print(
            self.color(
                f"🔴 Withdraw : Rs.{total_withdraw:,.2f}",
                self.RED
            )
        )

        print()

        deposit_percent = (
            total_deposit / total
        ) * 100

        withdraw_percent = (
            total_withdraw / total
        ) * 100

        print(
            self.color(
                f"Deposit  = {deposit_percent:.1f}%",
                self.GREEN
            )
        )

        print(
            self.color(
                f"Withdraw = {withdraw_percent:.1f}%",
                self.RED
            )
        )

        print()

        # -----------------------------------------------------
        # ONE HORIZONTAL SUMMARY BAR
        # -----------------------------------------------------

        width = 42

        deposit_length = int(
            (total_deposit / total) * width
        )

        withdraw_length = (
            width - deposit_length
        )

        if (
            total_deposit > 0
            and deposit_length == 0
        ):
            deposit_length = 1
            withdraw_length = width - 1

        if (
            total_withdraw > 0
            and withdraw_length == 0
        ):
            withdraw_length = 1
            deposit_length = width - 1

        print(
            self.color(
                "█" * deposit_length,
                self.GREEN
            )
            +
            self.color(
                "█" * withdraw_length,
                self.RED
            )
        )

        print(
            self.color(
                "Deposit",
                self.GREEN
            )
            +
            " " * 25
            +
            self.color(
                "Withdraw",
                self.RED
            )
        )

        print(
            self.color(
                "─" * 58,
                self.BLUE
            )
        )

    # =========================================================
    # CHANGE PIN
    # =========================================================

    def change_pin(self):

        old = input(
            "Enter Old PIN: "
        ).strip()

        if old != self.pin:

            print(
                "Wrong PIN!"
            )

            return

        new = input(
            "Enter New PIN (4 digits): "
        ).strip()

        while (
            len(new) != 4
            or not new.isdigit()
        ):

            print(
                "PIN must be exactly 4 digits!"
            )

            new = input(
                "Enter New PIN (4 digits): "
            ).strip()

        if new == self.pin:

            print(
                "New PIN cannot be the same as old PIN."
            )

            return

        if new in self.registry["used_pins"]:

            print(
                "This PIN is already used by another account!"
            )

            return

        if self.pin in self.registry["used_pins"]:

            self.registry["used_pins"].remove(
                self.pin
            )

        self.registry["used_pins"].append(
            new
        )

        username = self.name.lower()

        if username in self.registry["accounts"]:

            self.registry["accounts"][username][
                "PIN"
            ] = new

        self.pin = new

        self.account["PIN"] = new

        self.save_registry()

        print(
            "PIN changed successfully!"
        )

    # =========================================================
    # EXPORT BALANCE EXCEL
    # =========================================================

    def export_balance_excel(self):

        try:

            wb = Workbook()

            ws = wb.active

            ws.title = "Account Balance"

            ws.append([
                "ACCOUNT BALANCE REPORT",
                ""
            ])

            ws.merge_cells("A1:B1")

            title_fill = PatternFill(
                start_color="1F4E78",
                end_color="1F4E78",
                fill_type="solid"
            )

            title_font = Font(
                name="Arial",
                size=16,
                bold=True,
                color="FFFFFF"
            )

            border = Border(
                left=Side(
                    style="medium",
                    color="1F4E78"
                ),
                right=Side(
                    style="medium",
                    color="1F4E78"
                ),
                top=Side(
                    style="medium",
                    color="1F4E78"
                ),
                bottom=Side(
                    style="medium",
                    color="1F4E78"
                )
            )

            ws["A1"].fill = title_fill
            ws["A1"].font = title_font
            ws["A1"].alignment = Alignment(
                horizontal="center"
            )
            ws["A1"].border = border

            ws["B1"].fill = title_fill
            ws["B1"].border = border

            details = [

                (
                    "Account Holder",
                    self.name
                ),

                (
                    "Account ID",
                    self.account_id
                ),

                (
                    "Account Type",
                    self.account_info[0]
                ),

                (
                    "Current Balance",
                    f"Rs.{self.balance:,.2f}"
                ),

                (
                    "Generated On",
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                )
            ]

            for row, (
                label,
                value
            ) in enumerate(
                details,
                3
            ):

                label_cell = ws.cell(
                    row=row,
                    column=1,
                    value=label
                )

                value_cell = ws.cell(
                    row=row,
                    column=2,
                    value=value
                )

                label_cell.font = Font(
                    bold=True
                )

                for cell in (
                    label_cell,
                    value_cell
                ):

                    cell.border = border

                    cell.alignment = Alignment(
                        horizontal="center"
                    )

            ws.column_dimensions[
                "A"
            ].width = 25

            ws.column_dimensions[
                "B"
            ].width = 40

            filename = os.path.join(
                self.accounts_folder,
                f"{self.account_id}_balance.xlsx"
            )

            wb.save(filename)

            wb.close()

            print(
                f"\nBalance exported to Excel:\n{filename}"
            )

            return filename

        except Exception as e:

            print(
                f"Error exporting balance: {e}"
            )

            return None

    # =========================================================
    # EXPORT BALANCE WORD
    # =========================================================

    def export_balance_word(self):

        try:

            doc = Document()

            title = doc.add_heading(
                "ACCOUNT BALANCE REPORT",
                0
            )

            title.alignment = 1

            table = doc.add_table(
                rows=5,
                cols=2
            )

            table.style = "Table Grid"

            data = [

                (
                    "Account Holder",
                    self.name
                ),

                (
                    "Account ID",
                    self.account_id
                ),

                (
                    "Account Type",
                    self.account_info[0]
                ),

                (
                    "Current Balance",
                    f"Rs.{self.balance:,.2f}"
                ),

                (
                    "Generated On",
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                )
            ]

            for i, (
                label,
                value
            ) in enumerate(data):

                table.rows[i].cells[0].text = label

                table.rows[i].cells[1].text = value

                if label == "Current Balance":

                    for paragraph in (
                        table.rows[i]
                        .cells[1]
                        .paragraphs
                    ):

                        for run in paragraph.runs:

                            run.bold = True

                            run.font.size = Pt(14)

                            run.font.color.rgb = RGBColor(
                                0,
                                128,
                                0
                            )

            filename = os.path.join(
                self.accounts_folder,
                f"{self.account_id}_balance.docx"
            )

            doc.save(filename)

            print(
                f"Balance exported to Word:\n{filename}"
            )

            return filename

        except Exception as e:

            print(
                f"Error exporting balance: {e}"
            )

            return None

    # =========================================================
    # EXPORT HISTORY TO EXCEL
    # ONE PIE CHART + ALL TRANSACTIONS BELOW
    # =========================================================

    def export_history_excel(self):

        """
        Create a professional ATM transaction report.

        Layout:
        1. ATM Transaction History title
        2. Complete account information
        3. Transaction summary (deposits, withdrawals, current balance)
        4. Deposit vs Withdraw chart
        5. Complete transaction table

        The chart is generated as an image and embedded into Excel.
        This avoids the blank-chart problem that can happen with
        openpyxl chart objects in some Excel/LibreOffice previews.
        """

        chart_image = None

        try:

            # =====================================================
            # READ ALL TRANSACTIONS
            # =====================================================

            records = []

            if os.path.exists(self.excel_file):

                source_wb = load_workbook(
                    self.excel_file,
                    data_only=True
                )

                source_ws = source_wb.active

                for row in source_ws.iter_rows(
                    min_row=2,
                    values_only=True
                ):

                    if not row:
                        continue

                    date_value = row[0]
                    transaction_type = row[1]
                    amount_value = row[2]
                    balance_value = row[3]

                    if not transaction_type:
                        continue

                    transaction_type = str(
                        transaction_type
                    ).strip()

                    normalized_type = transaction_type.lower()

                    if normalized_type not in {
                        "initial balance",
                        "deposit",
                        "withdraw",
                        "withdrawal"
                    }:
                        continue

                    try:
                        amount_value = float(
                            amount_value or 0
                        )
                    except (TypeError, ValueError):
                        amount_value = 0.0

                    try:
                        balance_value = float(
                            balance_value or 0
                        )
                    except (TypeError, ValueError):
                        balance_value = 0.0

                    if normalized_type == "withdrawal":
                        transaction_type = "Withdraw"
                    elif normalized_type == "withdraw":
                        transaction_type = "Withdraw"
                    elif normalized_type == "deposit":
                        transaction_type = "Deposit"
                    else:
                        transaction_type = "Initial Balance"

                    records.append({
                        "date": date_value,
                        "transaction": transaction_type,
                        "amount": abs(amount_value),
                        "balance": balance_value
                    })

                source_wb.close()

            # =====================================================
            # CALCULATE SUMMARY
            # =====================================================

            total_deposit = sum(
                record["amount"]
                for record in records
                if record["transaction"] == "Deposit"
            )

            total_withdraw = sum(
                record["amount"]
                for record in records
                if record["transaction"] == "Withdraw"
            )

            deposit_count = sum(
                1
                for record in records
                if record["transaction"] == "Deposit"
            )

            withdraw_count = sum(
                1
                for record in records
                if record["transaction"] == "Withdraw"
            )

            initial_balance = 10000.0

            for record in records:
                if record["transaction"] == "Initial Balance":
                    initial_balance = record["amount"]
                    break

            # =====================================================
            # CREATE WORKBOOK
            # =====================================================

            wb = Workbook()

            ws = wb.active
            ws.title = "ATM Transaction History"

            # =====================================================
            # COLORS
            # =====================================================

            purple = "7030A0"
            dark_purple = "4B237A"
            white = "FFFFFF"

            green = "006100"
            light_green = "C6EFCE"

            red = "9C0006"
            light_red = "FFC7CE"

            gold = "7F6000"
            light_gold = "FFF2CC"

            light_gray = "F3F3F3"
            border_color = "B7B7B7"

            thin_side = Side(
                style="thin",
                color=border_color
            )

            medium_side = Side(
                style="medium",
                color=purple
            )

            thin_border = Border(
                left=thin_side,
                right=thin_side,
                top=thin_side,
                bottom=thin_side
            )

            medium_border = Border(
                left=medium_side,
                right=medium_side,
                top=medium_side,
                bottom=medium_side
            )

            # =====================================================
            # TITLE
            # =====================================================

            ws.merge_cells("A1:F1")

            ws["A1"] = "ATM TRANSACTION HISTORY"

            ws["A1"].font = Font(
                name="Arial",
                size=20,
                bold=True,
                color=white
            )

            ws["A1"].fill = PatternFill(
                fill_type="solid",
                fgColor=purple
            )

            ws["A1"].alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            ws.row_dimensions[1].height = 36

            # =====================================================
            # ACCOUNT INFORMATION
            # =====================================================

            ws.merge_cells("A3:F3")

            ws["A3"] = "ACCOUNT INFORMATION"

            ws["A3"].font = Font(
                name="Arial",
                size=14,
                bold=True,
                color=white
            )

            ws["A3"].fill = PatternFill(
                fill_type="solid",
                fgColor=dark_purple
            )

            ws["A3"].alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            account_details = [
                ("Account Holder", self.name),
                ("ATM / Account Number", self.account_id),
                ("Account Type", self.account_info[0]),
                ("Initial Balance", initial_balance),
                ("Current Balance", self.balance),
                (
                    "Generated On",
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                )
            ]

            detail_row = 4

            for label, value in account_details:

                ws.cell(
                    detail_row,
                    1,
                    label
                )

                ws.merge_cells(
                    start_row=detail_row,
                    start_column=2,
                    end_row=detail_row,
                    end_column=6
                )

                value_cell = ws.cell(
                    detail_row,
                    2,
                    value
                )

                label_cell = ws.cell(
                    detail_row,
                    1
                )

                label_cell.font = Font(
                    bold=True,
                    color=dark_purple
                )

                label_cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor=light_gray
                )

                value_cell.font = Font(
                    bold=True
                    if label in {
                        "Account Holder",
                        "ATM / Account Number",
                        "Current Balance"
                    }
                    else False
                )

                if label in {
                    "Initial Balance",
                    "Current Balance"
                }:
                    value_cell.number_format = (
                        '"Rs. " #,##0.00'
                    )

                for col in range(1, 7):

                    cell = ws.cell(
                        detail_row,
                        col
                    )

                    cell.border = thin_border
                    cell.alignment = Alignment(
                        vertical="center",
                        horizontal=(
                            "left"
                            if col == 1
                            else "center"
                        )
                    )

                detail_row += 1

            # Highlight current balance
            ws["B8"].fill = PatternFill(
                fill_type="solid",
                fgColor=light_gold
            )

            ws["B8"].font = Font(
                bold=True,
                size=13,
                color=gold
            )

            # =====================================================
            # TRANSACTION SUMMARY
            # =====================================================

            summary_title_row = 11

            ws.merge_cells(
                start_row=summary_title_row,
                start_column=1,
                end_row=summary_title_row,
                end_column=6
            )

            ws.cell(
                summary_title_row,
                1,
                "TRANSACTION SUMMARY"
            )

            ws.cell(
                summary_title_row,
                1
            ).font = Font(
                size=14,
                bold=True,
                color=white
            )

            ws.cell(
                summary_title_row,
                1
            ).fill = PatternFill(
                fill_type="solid",
                fgColor=purple
            )

            ws.cell(
                summary_title_row,
                1
            ).alignment = Alignment(
                horizontal="center"
            )

            summary_headers = [
                "Total Deposits",
                "Deposit Count",
                "Total Withdrawals",
                "Withdrawal Count",
                "Current Balance",
                "Net Movement"
            ]

            summary_values = [
                total_deposit,
                deposit_count,
                total_withdraw,
                withdraw_count,
                self.balance,
                total_deposit - total_withdraw
            ]

            for col, value in enumerate(
                summary_headers,
                1
            ):

                cell = ws.cell(
                    summary_title_row + 1,
                    col,
                    value
                )

                cell.font = Font(
                    bold=True,
                    color=white
                )

                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor=dark_purple
                )

                cell.border = thin_border

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True
                )

            for col, value in enumerate(
                summary_values,
                1
            ):

                cell = ws.cell(
                    summary_title_row + 2,
                    col,
                    value
                )

                cell.border = thin_border

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

                if col in {
                    1,
                    3,
                    5,
                    6
                }:
                    cell.number_format = (
                        '"Rs. " #,##0.00'
                    )

                cell.font = Font(
                    bold=True,
                    size=12
                )

            # Deposit
            ws.cell(
                summary_title_row + 2,
                1
            ).fill = PatternFill(
                fill_type="solid",
                fgColor=light_green
            )

            ws.cell(
                summary_title_row + 2,
                1
            ).font = Font(
                bold=True,
                color=green,
                size=12
            )

            # Withdrawal
            ws.cell(
                summary_title_row + 2,
                3
            ).fill = PatternFill(
                fill_type="solid",
                fgColor=light_red
            )

            ws.cell(
                summary_title_row + 2,
                3
            ).font = Font(
                bold=True,
                color=red,
                size=12
            )

            # Current balance
            ws.cell(
                summary_title_row + 2,
                5
            ).fill = PatternFill(
                fill_type="solid",
                fgColor=light_gold
            )

            ws.cell(
                summary_title_row + 2,
                5
            ).font = Font(
                bold=True,
                color=gold,
                size=12
            )

            # =====================================================
            # CREATE GUARANTEED-VISIBLE CHART IMAGE
            # =====================================================

            chart_image = os.path.join(
                self.accounts_folder,
                f"{self.account_id}_deposit_withdraw_chart.png"
            )

            labels = ["Deposits", "Withdrawals"]
            values = [
                total_deposit,
                total_withdraw
            ]

            plt.figure(
                figsize=(8, 4.8),
                dpi=150
            )

            if sum(values) > 0:

                plt.pie(
                    values,
                    labels=labels,
                    autopct="%1.1f%%",
                    startangle=90,
                    wedgeprops={
                        "width": 0.45
                    }
                )

                plt.title(
                    "Deposit vs Withdraw",
                    fontsize=16,
                    fontweight="bold"
                )

                plt.tight_layout()

            else:

                plt.text(
                    0.5,
                    0.5,
                    "No deposit or withdrawal data",
                    ha="center",
                    va="center",
                    fontsize=14
                )

                plt.axis("off")

            plt.savefig(
                chart_image,
                bbox_inches="tight"
            )

            plt.close()

            # =====================================================
            # CHART SECTION
            # =====================================================

            chart_title_row = 15

            ws.merge_cells(
                start_row=chart_title_row,
                start_column=1,
                end_row=chart_title_row,
                end_column=6
            )

            ws.cell(
                chart_title_row,
                1,
                "DEPOSIT vs WITHDRAWAL CHART"
            )

            ws.cell(
                chart_title_row,
                1
            ).font = Font(
                size=14,
                bold=True,
                color=white
            )

            ws.cell(
                chart_title_row,
                1
            ).fill = PatternFill(
                fill_type="solid",
                fgColor=dark_purple
            )

            ws.cell(
                chart_title_row,
                1
            ).alignment = Alignment(
                horizontal="center"
            )

            excel_chart_image = XLImage(
                chart_image
            )

            excel_chart_image.width = 650
            excel_chart_image.height = 380

            ws.add_image(
                excel_chart_image,
                "A16"
            )

            # Give the chart enough vertical space.
            for row_number in range(16, 38):
                ws.row_dimensions[row_number].height = 22

            # =====================================================
            # TRANSACTION TABLE
            # =====================================================

            table_title_row = 40

            ws.merge_cells(
                start_row=table_title_row,
                start_column=1,
                end_row=table_title_row,
                end_column=6
            )

            ws.cell(
                table_title_row,
                1,
                "ALL DEPOSIT & WITHDRAW TRANSACTIONS"
            )

            ws.cell(
                table_title_row,
                1
            ).font = Font(
                size=14,
                bold=True,
                color=white
            )

            ws.cell(
                table_title_row,
                1
            ).fill = PatternFill(
                fill_type="solid",
                fgColor=purple
            )

            ws.cell(
                table_title_row,
                1
            ).alignment = Alignment(
                horizontal="center"
            )

            header_row = table_title_row + 1

            headers = [
                "#",
                "Date",
                "Transaction",
                "Amount",
                "Balance",
                "Status"
            ]

            for col, header in enumerate(
                headers,
                1
            ):

                cell = ws.cell(
                    header_row,
                    col,
                    header
                )

                cell.font = Font(
                    bold=True,
                    color=white
                )

                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor=dark_purple
                )

                cell.border = medium_border

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            table_row = header_row + 1
            transaction_number = 1

            for record in records:

                transaction_type = record["transaction"]

                if transaction_type not in {
                    "Deposit",
                    "Withdraw"
                }:
                    continue

                ws.cell(
                    table_row,
                    1,
                    transaction_number
                )

                ws.cell(
                    table_row,
                    2,
                    record["date"]
                )

                ws.cell(
                    table_row,
                    3,
                    transaction_type
                )

                ws.cell(
                    table_row,
                    4,
                    record["amount"]
                )

                ws.cell(
                    table_row,
                    5,
                    record["balance"]
                )

                ws.cell(
                    table_row,
                    6,
                    "Credit"
                    if transaction_type == "Deposit"
                    else "Debit"
                )

                ws.cell(
                    table_row,
                    4
                ).number_format = (
                    '"Rs. " #,##0.00'
                )

                ws.cell(
                    table_row,
                    5
                ).number_format = (
                    '"Rs. " #,##0.00'
                )

                if transaction_type == "Deposit":

                    fill = PatternFill(
                        fill_type="solid",
                        fgColor=light_green
                    )

                    font_color = green

                else:

                    fill = PatternFill(
                        fill_type="solid",
                        fgColor=light_red
                    )

                    font_color = red

                for col in range(1, 7):

                    cell = ws.cell(
                        table_row,
                        col
                    )

                    cell.border = thin_border
                    cell.fill = fill

                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center"
                    )

                ws.cell(
                    table_row,
                    3
                ).font = Font(
                    bold=True,
                    color=font_color
                )

                ws.cell(
                    table_row,
                    4
                ).font = Font(
                    bold=True,
                    color=font_color
                )

                transaction_number += 1
                table_row += 1

            if transaction_number == 1:

                ws.merge_cells(
                    start_row=table_row,
                    start_column=1,
                    end_row=table_row,
                    end_column=6
                )

                ws.cell(
                    table_row,
                    1,
                    "No Deposit or Withdraw transactions found."
                )

                ws.cell(
                    table_row,
                    1
                ).alignment = Alignment(
                    horizontal="center"
                )

            # =====================================================
            # FINAL BALANCE
            # =====================================================

            current_row = table_row + 2

            ws.merge_cells(
                start_row=current_row,
                start_column=1,
                end_row=current_row,
                end_column=5
            )

            ws.cell(
                current_row,
                1,
                "CURRENT BALANCE"
            )

            ws.cell(
                current_row,
                1
            ).font = Font(
                bold=True,
                color=white
            )

            ws.cell(
                current_row,
                1
            ).fill = PatternFill(
                fill_type="solid",
                fgColor=dark_purple
            )

            ws.cell(
                current_row,
                1
            ).alignment = Alignment(
                horizontal="right"
            )

            ws.cell(
                current_row,
                6,
                self.balance
            )

            ws.cell(
                current_row,
                6
            ).number_format = (
                '"Rs. " #,##0.00'
            )

            ws.cell(
                current_row,
                6
            ).font = Font(
                bold=True,
                color=gold,
                size=12
            )

            ws.cell(
                current_row,
                6
            ).fill = PatternFill(
                fill_type="solid",
                fgColor=light_gold
            )

            for col in range(1, 7):

                ws.cell(
                    current_row,
                    col
                ).border = medium_border

            # =====================================================
            # COLUMN WIDTHS
            # =====================================================

            widths = {
                "A": 10,
                "B": 24,
                "C": 20,
                "D": 20,
                "E": 20,
                "F": 18
            }

            for column, width in widths.items():

                ws.column_dimensions[
                    column
                ].width = width

            # =====================================================
            # SHEET FORMATTING
            # =====================================================

            ws.freeze_panes = f"A{header_row + 1}"

            if transaction_number > 1:

                ws.auto_filter.ref = (
                    f"A{header_row}:"
                    f"F{table_row - 1}"
                )

            ws.sheet_view.showGridLines = False

            ws.page_setup.orientation = "landscape"
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0

            ws.sheet_properties.pageSetUpPr.fitToPage = True

            ws.print_area = (
                f"A1:F{current_row}"
            )

            # =====================================================
            # SECOND SHEET: RAW TRANSACTIONS
            # =====================================================

            raw_ws = wb.create_sheet(
                "Transactions"
            )

            raw_headers = [
                "Date",
                "Transaction",
                "Amount",
                "Balance"
            ]

            for col, header in enumerate(
                raw_headers,
                1
            ):

                cell = raw_ws.cell(
                    1,
                    col,
                    header
                )

                cell.font = Font(
                    bold=True,
                    color=white
                )

                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor=dark_purple
                )

                cell.border = thin_border

                cell.alignment = Alignment(
                    horizontal="center"
                )

            raw_row = 2

            for record in records:

                raw_ws.cell(
                    raw_row,
                    1,
                    record["date"]
                )

                raw_ws.cell(
                    raw_row,
                    2,
                    record["transaction"]
                )

                raw_ws.cell(
                    raw_row,
                    3,
                    record["amount"]
                )

                raw_ws.cell(
                    raw_row,
                    4,
                    record["balance"]
                )

                raw_ws.cell(
                    raw_row,
                    3
                ).number_format = (
                    '"Rs. " #,##0.00'
                )

                raw_ws.cell(
                    raw_row,
                    4
                ).number_format = (
                    '"Rs. " #,##0.00'
                )

                raw_row += 1

            for column in ["A", "B", "C", "D"]:
                raw_ws.column_dimensions[
                    column
                ].width = 24

            raw_ws.freeze_panes = "A2"

            if raw_row > 2:

                raw_ws.auto_filter.ref = (
                    f"A1:D{raw_row - 1}"
                )

            raw_ws.sheet_view.showGridLines = False

            # =====================================================
            # SAVE
            # =====================================================

            filename = os.path.join(
                self.accounts_folder,
                f"{self.account_id}_history.xlsx"
            )

            wb.save(filename)
            wb.close()

            print(
                self.color(
                    "\n📊 Professional Excel history exported successfully!",
                    self.GREEN
                )
            )

            print(
                self.color(
                    f"File: {filename}",
                    self.WHITE
                )
            )

            print(
                self.color(
                    "✓ Complete account information",
                    self.GREEN
                )
            )

            print(
                self.color(
                    "✓ Deposit / withdrawal summary",
                    self.GREEN
                )
            )

            print(
                self.color(
                    "✓ Visible Deposit vs Withdrawal chart",
                    self.GREEN
                )
            )

            print(
                self.color(
                    "✓ All transactions in a formatted table",
                    self.GREEN
                )
            )

            return filename

        except PermissionError:

            print(
                self.color(
                    "Cannot save Excel. "
                    "Please close the Excel file first.",
                    self.RED
                )
            )

            return None

        except Exception as e:

            print(
                self.color(
                    f"Error exporting history: {e}",
                    self.RED
                )
            )

            return None

        finally:

            # The PNG is only an intermediate chart file.
            # Keep it if Excel could not be created so the
            # user can inspect the chart separately.
            pass

    # =========================================================
    # EXPORT HISTORY WORD
    # =========================================================

    def export_history_word(self):

        try:

            doc = Document()

            title = doc.add_heading(
                "ATM Transaction History Report",
                0
            )

            title.alignment = 1

            doc.add_paragraph(
                f"Account Holder: {self.name}"
            )

            doc.add_paragraph(
                f"Account ID: {self.account_id}"
            )

            doc.add_paragraph(
                f"Account Type: {self.account_info[0]}"
            )

            doc.add_paragraph(
                "Generated On: " +
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            )

            doc.add_paragraph(
                "-" * 60
            )

            doc.add_paragraph(
                "Initial Balance: Rs.10,000.00"
            )

            for i, transaction in enumerate(
                self.transactions,
                1
            ):

                doc.add_paragraph(
                    f"{i}. {transaction}"
                )

            doc.add_paragraph(
                "-" * 60
            )

            doc.add_paragraph(
                f"Current Balance: "
                f"Rs.{self.balance:,.2f}"
            )

            filename = os.path.join(
                self.accounts_folder,
                f"{self.account_id}_history.docx"
            )

            doc.save(filename)

            print(
                f"History exported to Word:\n{filename}"
            )

            return filename

        except Exception as e:

            print(
                f"Error exporting history: {e}"
            )

            return None

    # =========================================================
    # ATM MENU
    # =========================================================

    def menu(self):

        while True:

            print(
                self.color(
                    "\n╔" + "═" * 48 + "╗",
                    self.CYAN
                )
            )

            print(
                self.color(
                    "║" +
                    "🏦 ATM MACHINE".center(48) +
                    "║",
                    self.CYAN
                )
            )

            print(
                self.color(
                    "╠" + "═" * 48 + "╣",
                    self.CYAN
                )
            )

            print(
                self.color(
                    f"║ 👋 Welcome   : {self.name}".ljust(49) +
                    "║",
                    self.GREEN
                )
            )

            print(
                self.color(
                    f"║ 🆔 Account ID: {self.account_id}".ljust(49) +
                    "║",
                    self.WHITE
                )
            )

            print(
                self.color(
                    f"║ 💰 Balance   : Rs.{self.balance:,.2f}".ljust(49) +
                    "║",
                    self.YELLOW
                )
            )

            print(
                self.color(
                    "╠" + "═" * 48 + "╣",
                    self.CYAN
                )
            )

            print(
                self.color(
                    "║ 1. 💵 Deposit".ljust(49) +
                    "║",
                    self.GREEN
                )
            )

            print(
                self.color(
                    "║ 2. 💳 Withdraw".ljust(49) +
                    "║",
                    self.RED
                )
            )

            print(
                self.color(
                    "║ 3. 💰 Balance".ljust(49) +
                    "║",
                    self.YELLOW
                )
            )

            print(
                self.color(
                    "║ 4. 📜 History + Chart".ljust(49) +
                    "║",
                    self.MAGENTA
                )
            )

            print(
                self.color(
                    "║ 5. 🔑 Change PIN".ljust(49) +
                    "║",
                    self.BLUE
                )
            )

            print(
                self.color(
                    "║ 6. 🚪 Logout".ljust(49) +
                    "║",
                    self.WHITE
                )
            )

            print(
                self.color(
                    "╚" + "═" * 48 + "╝",
                    self.CYAN
                )
            )

            choice = input(
                "Enter Choice: "
            ).strip()

            if choice == "1":

                self.deposit()

            elif choice == "2":

                self.withdraw()

            elif choice == "3":

                self.check_balance()

            elif choice == "4":

                self.history()

            elif choice == "5":

                self.change_pin()

            elif choice == "6":

                print(
                    "\nLogged out successfully!"
                )

                break

            else:

                print(
                    "Invalid choice! Please select 1-6."
                )

    # =========================================================
    # START
    # =========================================================

    def start(self):

        while True:

            print(
                "\n" + "=" * 55
            )

            print(
                "              WELCOME TO ATM SYSTEM"
            )

            print(
                "=" * 55
            )

            print(
                "1. Create New Account"
            )

            print(
                "2. Login"
            )

            print(
                "3. Exit"
            )

            print(
                "=" * 55
            )

            choice = input(
                "Enter Choice: "
            ).strip()

            if choice == "1":

                self.create_account()

            elif choice == "2":

                if self.login():

                    self.menu()

            elif choice == "3":

                print(
                    "\nThank You for Using ATM System!"
                )

                break

            else:

                print(
                    "Invalid choice! Please select 1-3."
                )


# =============================================================
# PROGRAM START
# =============================================================

if __name__ == "__main__":

    atm = ATM()

    atm.start()
