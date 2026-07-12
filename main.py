import base64
import json
import shutil
from pathlib import Path
from enum import Enum

import fitz
from openai import OpenAI
from pydantic import BaseModel, Field
from typing import Optional, List

client = OpenAI()

SYSTEM_PROMPT = Path("prompts/invoice_extraction.txt").read_text(
    encoding="utf-8"
)

class Currency(str, Enum):
    EUR = "EUR"
    PLN = "PLN"
    GBP = "GBP"

class PaymentMethod(str, Enum):
    BANK_TRANSFER   = "BANK_TRANSFER"
    CASH            = "CASH"

class ServicePeriod(BaseModel):
    start_date: Optional[str] = None
    end_date: Optional[str] = None
    start_month: Optional[str] = None
    end_month: Optional[str] = None
    raw_text: Optional[str] = None

class InvoiceItem(BaseModel):
    number:             Optional[int]
    name_of_service:    Optional[str]
    place_of_service:   Optional[str]
    service_periods:    List[ServicePeriod] = []
    quantity:           Optional[float]
    unit:               Optional[str]
    unit_price:         Optional[float]
    net_amount:         Optional[float]
    vat_rate:           Optional[float]
    vat_amount:         Optional[float]
    gross_amount:       Optional[float]

class InvoiceBase(BaseModel):
    number:             Optional[str]
    issue_date:         Optional[str]
    sale_date:          Optional[str]
    service_periods:    List[ServicePeriod] = []
    place_of_service:   Optional[str]
    seller_name:        Optional[str]
    seller_address:     Optional[str]
    seller_tax_id:      Optional[str]
    buyer_name:         Optional[str]
    buyer_address:      Optional[str]
    buyer_tax_id:       Optional[str]
    currency:           Optional[Currency]
    payment_method:     Optional[PaymentMethod]
    total_net:          Optional[float]
    total_vat:          Optional[float]
    total_gross:        Optional[float]

class Invoice(InvoiceBase):
    items: List[InvoiceItem]

class ReasonedInvoiceItem(InvoiceItem):
    confidence: float       = Field(ge=0, le=1)
    warnings:   List[str]   = []

class ReasonedInvoice(InvoiceBase):
    items:              List[ReasonedInvoiceItem]
    total_confidence:   float       = Field(ge=0, le=1)
    warnings:           List[str]   = []

def pdf_to_images(pdf_path: Path, out_dir: Path) -> list[Path]:
    doc = fitz.open(pdf_path)
    image_paths = []

    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_path = out_dir / f"{pdf_path.stem}_page_{i+1}.png"
        pix.save(img_path)
        image_paths.append(img_path)

    return image_paths

def image_to_base64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("utf-8")

def extract_invoice_from_images(image_paths: list[Path]) -> ReasonedInvoice:
    content = [
        {
            "type": "input_text",
            "text": (
                "Extract invoice data from these invoice pages. "
                "Return structured data only. One invoice line = one item."
            ),
        }
    ]

    for img_path in image_paths:
        b64 = image_to_base64(img_path)
        content.append({
            "type": "input_image",
            "image_url": f"data:image/png;base64,{b64}",
        })

    response = client.responses.parse(
        model="gpt-5.4",
        input=[
            {
                "role": "system",
                "content": SYSTEM_PROMPT,
            },
            {
                "role": "user",
                "content": content,
            },
        ],
        text_format=ReasonedInvoice,
    )

    return response.output_parsed


# def validate_invoice(inv: ReasonedInvoice) -> list[str]:
#     warnings = []

#     for idx, item in enumerate(inv.items, start=1):
#         if item.quantity and item.unit_price and item.net_amount:
#             expected = round(item.quantity * item.unit_price, 2)
#             actual = round(item.net_amount, 2)

#             if abs(expected - actual) > 0.05:
#                 warnings.append(
#                     f"Item {idx}: quantity * unit_price != net_amount "
#                     f"({expected} != {actual})"
#                 )

#         if item.net_amount is not None and item.vat_amount is not None and item.gross_amount is not None:
#             expected = round(item.net_amount + item.vat_amount, 2)
#             actual = round(item.gross_amount, 2)

#             if abs(expected - actual) > 0.05:
#                 warnings.append(
#                     f"Item {idx}: net + vat != gross "
#                     f"({expected} != {actual})"
#                 )

#         if item.confidence < 0.8:
#             warnings.append(f"Item {idx}: low confidence")

#     return warnings


def main():
    input_dir = Path("input")
    output_dir = Path("output")
    tmp_dir = Path("tmp")

    output_dir.mkdir(exist_ok=True)

    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
    tmp_dir.mkdir()

    try:
        for pdf_path in input_dir.glob("*.pdf"):
            print(f"Processing: {pdf_path.name}")

            image_paths = pdf_to_images(pdf_path, tmp_dir)
            invoice = extract_invoice_from_images(image_paths)
            # validation_warnings = validate_invoice(result)

            # print(json.dumps(invoice.model_dump(), ensure_ascii=False, indent=2))

            output_path = output_dir / f"{pdf_path.stem}.json"

            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(invoice.model_dump(mode="json"), f, ensure_ascii=False, indent=2)

            print(f"Output saved to: {output_path}")

            # if validation_warnings:
            #     print("\nVALIDATION WARNINGS:")
            #     for w in validation_warnings:
            #         print("-", w)

    finally:
        if tmp_dir.exists():
            shutil.rmtree(tmp_dir)

if __name__ == "__main__":
    main()
