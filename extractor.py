import base64
from pathlib import Path
from typing import Any, Dict, Optional, Tuple
import pypdf
import pymupdf
import config

class PDFFormExtractor:
    def __init__(self, pdf_path:Path):
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.is_file():
            raise FileNotFoundError(f"PDF File not Found in: {self.pdf_path}")
    def extract_all(self) -> Dict[str, Any]:
        record : Dict[str, Any] = {config.SYSTEM_COLUMNS["source_file"]:self.pdf_path.name}
        form_fields = self._extract_acroforms_fields()
        for raw_key, raw_value in form_fields.items():
            column_label = self._resolve_column_name(raw_key)
            record[column_label] = self._normalize_value(raw_value)
        sig_detected,sig_path, sig_base64_str = self._extract_embedded_signature()
        record[config.SYSTEM_COLUMNS["signature_status"]] = "YES" if sig_detected else "NO"
        if config.SIGNATURE_CAPTURE_MODE in ("hyperlink", "both"):
            record[config.SYSTEM_COLUMNS["signature_link"]] = str(sig_path) if sig_path else "N/A"
        if config.SIGNATURE_CAPTURE_MODE in ("both", "base64"):
            record[config.SYSTEM_COLUMNS["signature_base64"]] = sig_base64_str if sig_base64_str else "N/A"
        return record

    def _extract_acroforms_fields(self) -> Dict[str, Any]:
        field_data: Dict[str, Any] = {}
        try:
            reader = pypdf.PdfReader(str(self.pdf_path))
            raw_fields = reader.get_fields() or {}
            for field_name, field_info in raw_fields.items():
                field_data[field_name] = field_info.get("/V","")
        except Exception as e:
            print(f"[WARNING] Forms Reading in {self.pdf_path.name} Failed: {e}")
        return field_data

    @staticmethod
    def _resolve_column_name(raw_key: str) -> str:
        if hasattr(config, "FIELD_MAPPING") and raw_key in config.FIELD_MAPPING:
            return config.FIELD_MAPPING[raw_key]
        return raw_key.replace("_"," ").strip().title()

    @staticmethod
    def _normalize_value(value:Any) -> str:
        if value is None:
            return ""
        value_str = str(value).strip()
        if value_str in ("/Yes","/YES", "1", "Yes","/True", "True", "TRUE", "/true", "true"):
            return "Selected"
        if value_str in ("/Off", "/OFF", "0", "NO", "off", "FALSE", "False", "false", "/FALSE", "/false", "/False", "None", "NONE",""):
            return "Not Selected"
        if value_str.startswith("/") and len(value_str) > 1:
            return value_str[1:]
        return value_str
    
    def _extract_embedded_signature(self) -> Tuple[bool, Optional[Path],Optional[str]]:
        sig_detected = False
        sig_hyperlink: Optional[Path] = None
        sig_base64_str: Optional[str] = None

        try:
            doc = pymupdf.open(str(self.pdf_path))
            if config.SIGNATURE_CAPTURE_MODE in ("hyperlink", "both"):
                config.SIGNATURES_DIR.mkdir(parents=True, exist_ok=True)
            for page_index in range(len(doc)):
                page = doc[page_index]
                image_list = page.get_images(full=True)
                if image_list:
                    xref = image_list[0][0]
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    ext = base_image["ext"]
                    sig_detected = True
                    if config.SIGNATURE_CAPTURE_MODE in ("hyperlink", "both"):
                        stem = self.pdf_path.stem
                        dest_file = config.SIGNATURES_DIR / f"{stem}_signature.{ext}"
                        dest_file.write_bytes(image_bytes)
                        sig_hyperlink = dest_file
                    if config.SIGNATURE_CAPTURE_MODE in ("both", "base64"):
                        sig_base64_str = base64.b64encode(image_bytes).decode("utf-8")
                    break
            doc.close()
        except Exception as e:
            print(f"[WARNING] Signature Extraction in {self.pdf_path.name} Failed: {e}")
        return sig_detected, sig_hyperlink, sig_base64_str
