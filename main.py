import base64
import json
import time
import shutil
from pathlib import Path
from enum import Enum

from rich.console import Console
from rich.panel import Panel
from rich.progress import track
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn, TimeElapsedColumn

import fitz
from openai import OpenAI
from pydantic import BaseModel, Field
from pydantic import ValidationError
from typing import Optional, List

client = OpenAI()
console = Console()

SYSTEM_PROMPT = Path("prompts/invoice_extraction.txt").read_text(
    encoding="utf-8"
)

# Adres w jezy ku kraju z ktorego pochodzi
# Dokladniejsza ekstrakcja nazw
# Usuwanie jakichs prefixow przy adresach typu str., ul. itd.
# Dokladniejsza ekstrakcja ulic/separacja od imion
# Nie uzywanie daty wykonania uslugi
# Poprawienie logicznego rozumienia czym jest invoice item
# Poprawa kodu pocztowego, moze 5 cyfr po prostu

# 2
# Dodatkowe slowka w nazwie ulicy
# Lepsze wyciananie nazwy
# Bank trasnfer na podstawie numeru konta, albo terminu platnosci, bo gotowka to odrazu

# 3
# Nazwa robot wygenerowana z merge'a
# Zachowywanie literek w numerze budynku
# Nie przerabiaj nazw miejscowosci nawet jesli wydaje sie ze jest zla
# Format dat

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

class Address(BaseModel):
    street:             Optional[str] = None
    building_number:    Optional[str] = None
    postal_code:        Optional[str] = None
    city:               Optional[str] = None
    country:            Optional[str] = None

class InvoiceItem(BaseModel):
    number:             Optional[int] = None
    name_of_service:    Optional[str] = None
    place_of_service:   Optional[Address] = None
    service_periods:    List[ServicePeriod] = []
    net_amount:         Optional[float] = None
    vat_amount:         Optional[float] = None
    gross_amount:       Optional[float] = None

class InvoiceBase(BaseModel):
    number:             Optional[str]
    issue_date:         Optional[str]
    seller_name:        Optional[str]
    seller_address:     Optional[Address]
    seller_tax_id:      Optional[str]
    buyer_name:         Optional[str]
    buyer_address:      Optional[Address]
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

def extract_invoice_from_images(image_paths: list[Path]):
    content = [
        {
            "type": "input_text",
            "text": (
                "Extract invoice data from these invoice pages. "
                "Return structured data only."
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

    usage = {
        "input_tokens": response.usage.input_tokens,
        "output_tokens": response.usage.output_tokens,
        "total_tokens": response.usage.total_tokens,
    }

    return response.output_parsed, usage


def main():
    invoices_dir = Path("invoices")
    inv_ctx_dir = Path("invoices_context")
    tmp_dir = Path("tmp")

    console.print(
        Panel(
f"""[bold]PDF Context Reasoning[/bold]

[cyan]Parameters:[/cyan]
• Invoices directory:   {invoices_dir}
• Context directory:    {inv_ctx_dir}""",
            title="Configuration",
        )
    )

    console.print()

    inv_ctx_dir.mkdir(exist_ok=True)

    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
    tmp_dir.mkdir()

    if not invoices_dir.exists():
        console.print(
            f"[red]✗ No invoices directory found: {invoices_dir}[/red]"
        )
        return

    inv_files = list(invoices_dir.iterdir())

    if any(
        file.is_file() and file.suffix.lower() != ".pdf"
        for file in inv_files
    ):
        console.print(
            f"[red]✗ Directory contains non-PDF files[/red]"
        )
        return

    if len(inv_files) == 0:
        console.print(
            f"[yellow]! Invoice directory is empty[/yellow]"
        )
        return

    file_form = "file" if len(inv_files) == 1 else "files"
    console.print(
        f"[green]✓ Found {len(inv_files)} PDF {file_form}[/green]\n"
    )

    results = []
    global_start = time.time()

    try:
        # Konfiguracja zaawansowanego paska postępu
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TaskProgressColumn(),
            TextColumn("•"),
            TimeElapsedColumn(), # Całkowity czas od startu
            TextColumn("•"),
            TextColumn("[cyan]{task.fields[avg_time]}[/cyan]"), # Nasze własne pole ze średnią
            console=console
        ) as progress:
            
            task = progress.add_task(
                "Processing invoices...", 
                total=len(inv_files),
                avg_time="Avg: -- s/inv"
            )

            for i, pdf in enumerate(inv_files, start=1):
                result = {
                    "file": pdf.name,
                    "status": None,
                    "error": None,
                    "time": None,
                    "tokens": None,
                }
                item_start = time.time()

                try:
                    image_paths = pdf_to_images(pdf, tmp_dir)
                    invoice, usage = extract_invoice_from_images(image_paths)
                    result["tokens"] = usage

                    output_path = inv_ctx_dir / f"{pdf.stem}.json"

                    with open(output_path, "w", encoding="utf-8") as f:
                        json.dump(invoice.model_dump(mode="json"), f, ensure_ascii=False, indent=2)

                    result["status"] = "SUCCESS"

                except json.JSONDecodeError as e:
                    result["status"] = "FAILED"
                    result["error"] = (
                        f"Invalid JSON response: {e.msg} "
                        f"(line {e.lineno}, column {e.colno})"
                    )

                except ValidationError as e:
                    result["status"] = "FAILED"
                    errors = []
                    for error in e.errors():
                        location = ".".join(
                            str(x) for x in error["loc"]
                        )
                        errors.append(
                            f"{location}: {error['msg']}"
                        )

                    result["error"] = (
                        "Schema validation failed: "
                        + "; ".join(errors)
                    )

                except FileNotFoundError as e:
                    result["status"] = "FAILED"
                    result["error"] = (
                        f"File not found: {e.filename}"
                    )

                except PermissionError as e:
                    result["status"] = "FAILED"
                    result["error"] = (
                        f"Permission denied: {e.filename}"
                    )

                except Exception as e:
                    result["status"] = "FAILED"
                    result["error"] = (
                        f"Unexpected {type(e).__name__}: {str(e)}"
                    )

                finally:
                    # Czas dla pojedynczej faktury
                    item_time = time.time() - item_start
                    result["time"] = round(item_time, 2)
                    results.append(result)

                    # Obliczanie średniego czasu po wykonaniu i-tej faktury
                    total_elapsed = time.time() - global_start
                    avg_time_val = total_elapsed / i

                    # Aktualizacja paska (przesunięcie o 1 do przodu i nadpisanie średniej)
                    progress.update(
                        task, 
                        advance=1, 
                        avg_time=f"Avg: {avg_time_val:.1f} s/inv"
                    )

    finally:
        if tmp_dir.exists():
            shutil.rmtree(tmp_dir)

    table = Table(
        title="Processing Summary"
    )

    table.add_column("Invoice")
    table.add_column("Status")
    table.add_column("Time")
    table.add_column("Tokens")
    table.add_column("Error")

    for r in results:
        if r["status"] == "SUCCESS":
            status = "[green]✓ SUCCESS[/green]"
        else:
            status = "[red]✗ FAILED[/red]"

        tokens = "-"

        if r["tokens"]:
            tokens = (
                f'{r["tokens"]["total_tokens"]} '
                f'(↑{r["tokens"]["input_tokens"]} '
                f'↓{r["tokens"]["output_tokens"]})'
            )

        table.add_row(
            r["file"],
            status,
            f'{r["time"]}s',
            tokens,
            r["error"] or "-"
        )

    console.print()
    console.print(table)
    
    failed = [
        r for r in results
        if r["status"] == "FAILED"
    ]

    if failed:
        console.print(
            Panel(
                f"""[red]Failed invoices: {len(failed)}[/red]""",
                title="Errors"
            )
        )
    else:
        console.print(
            Panel(
                "[green]All invoices processed successfully![/green]",
                title="Done"
            )
        )

if __name__ == "__main__":
    main()
