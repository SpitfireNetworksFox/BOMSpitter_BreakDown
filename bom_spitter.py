# bom_writer.py
from __future__ import annotations
import os
from typing import List, Dict, Optional
import pandas as pd


# Usage :  python fire_a_quote.py --excel ".\Consilio - Q-586929-1.xlsx"

class BOMWriter:
    """
    Standalone Excel BOM writer that mirrors your sample's layout/formatting.
    - Uses pandas + xlsxwriter only (no DB).
    - Builds formula columns (Extended, Your Price, Our Cost, Margin) per row.
    - Adds a totals footer row with the same formulas you showed.
    - Applies the same fonts, widths, header styling, conditional left column,
      freeze panes, and optional logo insertion.
    """

    def __init__(self, font_name: str = "Futura Medium"):
        self.font_name = font_name

    # ----- Helpers that mirror your original formulas (rownum is 1-based in Excel) -----
    def _f_extended(self, rownum: int) -> str:
        # =((B{row}*G{row})*H{row})
        return f"=((B{rownum}*G{rownum})*H{rownum})"

    def _f_your_price(self, rownum: int) -> str:
        # =(I{row}-(I{row}*J{row}))
        return f"=(I{rownum}-(I{rownum}*J{rownum}))"

    def _f_our_cost(self, rownum: int) -> str:
        # =(I{row}-(I{row}*L{row}))
        return f"=(I{rownum}-(I{rownum}*L{rownum}))"

    def _f_margin(self, rownum: int) -> str:
        # =IFERROR(((K{row}-M{row})/K{row}),0)
        return f"=IFERROR(((K{rownum}-M{rownum})/K{rownum}),0)"

    # ---- Public API --------------------------------------------------------------------
    def build_dataframe(self, items: List[Dict]) -> pd.DataFrame:
        """
        Accepts 'items' with at least:
          qty, type, sku, pt_sku, description, price, term (months), sell_disc (% as 0-100), buy_disc (% as 0-100)
        Builds a DataFrame with the same columns as your sample (including formula columns).
        """
        rows = []
        # Excel row numbers start at 1; header is row 1; first data row is row 2
        for idx, it in enumerate(items, start=2):
            # Safe gets with defaults
            qty = it.get("qty", 1)
            prod_type = it.get("type", "")
            sku = it.get("sku", "")
            pt_sku = it.get("pt_sku", "")
            desc = it.get("description", "")
            price = it.get("price", 0.0)
            term_years = (int(it.get("term", 12)) or 12) / 12  # months → years
            sell_disc = float(it.get("sell_disc", 0.0)) / 100.0
            buy_disc = float(it.get("buy_disc", 0.0)) / 100.0

            row = {
                " ": "",
                "QTY": qty,
                "Product Type": prod_type,
                "SKU": sku,
                "Partner SKU": pt_sku,
                "Description": desc,
                "Price": price,
                "Term": term_years,
                "Extended": self._f_extended(idx),
                "Discount": sell_disc,
                "Your Price": self._f_your_price(idx),
                "Buy Discount": buy_disc,
                "Our Cost": self._f_our_cost(idx),
                "Margin": self._f_margin(idx),
            }
            rows.append(row)

        df = pd.DataFrame(rows)
        # Remove any "Unnamed: ..." columns if a caller passes a pre-built frame through
        df.drop(df.filter(regex="^Unname", axis=1), axis=1, errors="ignore", inplace=True)
        return df

    def write_bom(
        self,
        df: pd.DataFrame,
        output_path: str,
        sheet_name: str = "BOM",
        logo_path: Optional[str] = None,
    ) -> str:
        """
        Writes df to an Excel file matching your sample's layout.
        - output_path: file path like "generated/MySheet.xlsx"
        - sheet_name: trimmed to Excel's 31-char limit
        - logo_path: optional path to an image (PNG/JPG) to place at top-left
        Returns the final output_path.
        """
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

        # Excel sheet name limit (31 chars)
        sname = sheet_name if len(sheet_name) <= 31 else f"{sheet_name[:29]}.."

        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sname, freeze_panes=(1, 0))
            workbook  = writer.book
            worksheet = writer.sheets[sname]

            # ----- Formats (mirroring your sample) -----
            text_format = workbook.add_format({"font_name": self.font_name, "text_wrap": False})
            wrap_format = workbook.add_format({"font_name": self.font_name, "text_wrap": True})
            money_format = workbook.add_format({"font_name": self.font_name, "num_format": "$#,##0.00", "align": "right"})
            percent_format = workbook.add_format({"font_name": self.font_name, "num_format": "0%"})
            qty_format = workbook.add_format({"font_name": self.font_name, "align": "Center"})

            left_col_format = workbook.add_format({
                "font_name": self.font_name, "bg_color": "#404040", "font_color": "#ffffff",
                "bold": True, "align": "center", "border": 1,
            })

            header_format = workbook.add_format({
                "font_name": self.font_name, "bold": True, "text_wrap": True, "valign": "center",
                "align": "center", "fg_color": "#404040", "border": 1, "font_color": "#ffffff",
            })

            footer_money_format = workbook.add_format({
                "font_name": self.font_name, "num_format": "$#,##0.00", "align": "right",
                "fg_color": "#404040", "bold": True, "border": 1, "font_color": "#ffffff",
            })

            footer_percent_format = workbook.add_format({
                "font_name": self.font_name, "num_format": "0%", "align": "right",
                "fg_color": "#404040", "bold": True, "border": 1, "font_color": "#ffffff",
            })

            # ----- Column widths / formats (A:ZZ baseline then specifics) -----
            worksheet.set_column("A:ZZ", 7, text_format)
            worksheet.set_column("A:A", 7)                 # " "
            worksheet.set_column("B:B", 6, qty_format)     # QTY
            worksheet.set_column("C:C", 14)                # Product Type
            worksheet.set_column("D:D", 30)                # SKU
            worksheet.set_column("E:E", 30)                # Partner SKU
            worksheet.set_column("F:F", 55, wrap_format)   # Description
            worksheet.set_column("G:G", 18, money_format)  # Price (List Price)
            worksheet.set_column("H:H", 7)                 # Term
            worksheet.set_column("I:I", 20, money_format)  # Your Price deps; here Extended/Your Price share columns
            worksheet.set_column("J:J", 10, percent_format)
            worksheet.set_column("K:K", 20, money_format)
            worksheet.set_column("L:L", 10, percent_format)
            worksheet.set_column("M:M", 20, money_format)
            worksheet.set_column("N:N", 8,  percent_format)

            # Header row tweaks + optional logo
            worksheet.set_row(0, 100)
            if logo_path and os.path.exists(logo_path):
                worksheet.insert_image(0, 0, logo_path, {"x_scale": 0.3, "y_scale": 0.3, "x_offset": 10, "y_offset": 10})

            # Rewrite header row with formatting and rename "Price" → "List Price"
            columns = list(df.columns)
            for col_idx, val in enumerate(columns):
                display_val = "List Price" if val == "Price" else val
                worksheet.write(0, col_idx, display_val, header_format)

            # Left-most banding (conditional on not blank)
            # Compute last data row index (0-based) and Excel row numbers
            # dim_rowmax is the highest row index that contains data (0-based)
            last_data_row_0based = worksheet.dim_rowmax  # includes header row
            # Apply conditional format from row 2 (index 1) through the last data row.
            worksheet.conditional_format(1, 0, last_data_row_0based, 0,
                                         {"type": "no_blanks", "format": left_col_format})

            # Footer (Totals) one row after the last data row
            last_row_excel = last_data_row_0based + 1  # Convert to 1-based Excel index
            totals_row_excel = last_row_excel + 1

            # Write Totals line and blanks with header_format to create the dark band
            worksheet.write_string(f"A{totals_row_excel}", "Totals", header_format)
            for col_letter in ["B","C","D","E","F","G","H"]:
                worksheet.write_blank(f"{col_letter}{totals_row_excel}", "", header_format)

            # Footer formulas (match your sample)
            worksheet.write_formula(f"I{totals_row_excel}",
                                    f"=SUM(I2:I{last_row_excel})",
                                    footer_money_format)

            worksheet.write_formula(f"J{totals_row_excel}",
                                    f"=SUM((I{totals_row_excel}-K{totals_row_excel})/I{totals_row_excel})",
                                    footer_percent_format)

            worksheet.write_formula(f"K{totals_row_excel}",
                                    f"=SUM(K2:K{last_row_excel})",
                                    footer_money_format)

            worksheet.write_formula(f"L{totals_row_excel}",
                                    f"=SUM((I{totals_row_excel}-M{totals_row_excel})/I{totals_row_excel})",
                                    footer_percent_format)

            worksheet.write_formula(f"M{totals_row_excel}",
                                    f"=SUM(M2:M{last_row_excel})",
                                    footer_money_format)

            worksheet.write_formula(f"N{totals_row_excel}",
                                    f"=SUM((K{totals_row_excel}-M{totals_row_excel})/K{totals_row_excel})",
                                    footer_percent_format)

            # Save
            # (context manager closes the writer)
        return output_path


# -------------------- Example usage --------------------
if __name__ == "__main__":
    items = [
        {
            "qty": 2,
            "type": "Hardware",
            "sku": "NT-EDGE-1000",
            "pt_sku": "NT-EDGE-1000-HW-AC",
            "description": "Edge gateway appliance with dual AC PSU",
            "price": 3499.00,
            "term": 12,          # months
            "sell_disc": 10.0,   # % discount for 'Your Price'
            "buy_disc": 5.0,     # % discount for 'Our Cost'
        },
        {
            "qty": 50,
            "type": "Subscription",
            "sku": "TR-SWBSUB-1YR",
            "pt_sku": "TR-SWBSUB",
            "description": "Threat Response Software Subscription (1 year)",
            "price": 49.95,
            "term": 12,
            "sell_disc": 15.0,
            "buy_disc": 8.0,
        },
        {
            "qty": 4,
            "type": "Transceiver",
            "sku": "SFP-1G-SX",
            "pt_sku": "SFP-1G-SX",
            "description": "1G SX SFP transceiver, 850nm, 550m",
            "price": 79.00,
            "term": 12,
            "sell_disc": 0.0,
            "buy_disc": 0.0,
        },
    ]

    writer = BOMWriter(font_name="Futura Medium")
    df = writer.build_dataframe(items)
    out_path = writer.write_bom(
        df=df,
        output_path="generated/MyDeal.xlsx",
        sheet_name="My Customer Deal",
        logo_path=None,  # e.g., "assets/logo.png"
    )
    print(f"Wrote: {out_path}")
