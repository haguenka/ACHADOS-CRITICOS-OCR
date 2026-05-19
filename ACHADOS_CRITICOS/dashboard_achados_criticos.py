#!/usr/bin/env python3
"""
Dashboard Moderno para Análise de Achados Críticos
Interface UI/UX Dark Mode com Streamlit + Plotly
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import io
import base64
import os
import re
import shutil
import smtplib
import unicodedata
import textwrap
from difflib import SequenceMatcher
from email.message import EmailMessage
from pathlib import Path
from PIL import ImageFilter
from PIL import ImageDraw

# Configuração da página
st.set_page_config(
    page_title="CDI - Achados Críticos Dashboard",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado para dark mode
st.markdown("""
<style>
    /* Dark theme principal */
    .main > div {
        padding-top: 2rem;
    }

    /* Cards com estilo glassmorphism */
    .metric-card {
        background: linear-gradient(145deg, rgba(255, 255, 255, 0.1), rgba(255, 255, 255, 0.05));
        backdrop-filter: blur(10px);
        border-radius: 15px;
        padding: 20px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        margin: 10px 0;
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.37);
    }

    /* Header personalizado */
    .main-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 30px;
        text-align: center;
    }

    .main-header h1 {
        color: white;
        margin: 0;
        font-size: 2.5rem;
        font-weight: 700;
    }

    .main-header p {
        color: rgba(255, 255, 255, 0.9);
        margin: 10px 0 0 0;
        font-size: 1.1rem;
    }

    /* Sidebar customizada */
    .css-1d391kg {
        background: linear-gradient(180deg, #2C3E50 0%, #34495E 100%);
    }

    /* Métricas destacadas */
    .big-metric {
        text-align: center;
        padding: 20px;
        background: linear-gradient(145deg, rgba(52, 152, 219, 0.2), rgba(155, 89, 182, 0.2));
        border-radius: 15px;
        margin: 10px 0;
    }

    .big-metric h2 {
        font-size: 3rem;
        margin: 0;
        font-weight: 700;
    }

    .big-metric p {
        font-size: 1.2rem;
        margin: 5px 0 0 0;
        opacity: 0.8;
    }

    /* Status indicators */
    .status-success { color: #27AE60; }
    .status-warning { color: #F39C12; }
    .status-danger { color: #E74C3C; }

    /* Upload area */
    .upload-area {
        border: 2px dashed rgba(255, 255, 255, 0.3);
        border-radius: 10px;
        padding: 40px;
        text-align: center;
        margin: 20px 0;
    }
</style>
""", unsafe_allow_html=True)

def _env_bool(name, default=True):
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() not in {"0", "false", "nao", "não", "no", "off"}


def _env_int(name, default):
    try:
        return int(os.environ.get(name, default))
    except (TypeError, ValueError):
        return default


# Configuracao do servidor SMTP por variaveis de ambiente.
ADMIN_SMTP_CONFIG = {
    "host": os.environ.get("SMTP_HOST", ""),
    "port": _env_int("SMTP_PORT", 587),
    "use_tls": _env_bool("SMTP_USE_TLS", True),
    "username": os.environ.get("SMTP_USERNAME", ""),
    "password": os.environ.get("SMTP_PASSWORD", ""),
    "sender_email": os.environ.get("SMTP_SENDER_EMAIL", os.environ.get("SMTP_USERNAME", "")),
    "sender_name": os.environ.get("SMTP_SENDER_NAME", "CDI - Achados Criticos"),
    "tesseract_cmd": os.environ.get("TESSERACT_CMD", "/usr/bin/tesseract"),
}

# Regioes relativas ao dialogo "Resultado Critico".
RIS_DIALOG_FIELD_REGIONS = [
    {"field": "Resultado Crítico", "box": (0.31, 0.13, 0.70, 0.21), "multiline": False},
    {"field": "Contato", "box": (0.20, 0.25, 0.59, 0.34), "multiline": False},
    {"field": "Contato com (Sucesso)", "box": (0.80, 0.245, 0.965, 0.335), "multiline": False},
    {"field": "Achado Crítico", "box": (0.18, 0.40, 0.98, 0.49), "multiline": False},
    {"field": "Data e Hora", "box": (0.19, 0.48, 0.59, 0.58), "multiline": False},
    {"field": "Observações", "box": (0.18, 0.70, 0.995, 0.96), "multiline": True},
]

RIS_DIAGNOSIS_RELATIVE_BOX = (-0.24, -0.205, 0.24, -0.08)

RIS_FIELD_OCR_RULES = {
    "Diagnóstico": {
        "psm": 7,
        "scale": 5,
        "whitelist": "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzÀÁÂÃÉÊÍÓÔÕÚÇàáâãéêíóôõúç, -:",
    },
    "Resultado Crítico": {
        "psm": 7,
        "scale": 5,
        "whitelist": "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzÀÁÂÃÉÊÍÓÔÕÚÇàáâãéêíóôõúç ",
    },
    "Contato": {
        "psm": 7,
        "scale": 5,
        "whitelist": "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzÀÁÂÃÉÊÍÓÔÕÚÇàáâãéêíóôõúç ",
    },
    "Contato com (Sucesso)": {
        "psm": 8,
        "scale": 10,
        "whitelist": "SsIiMmNnAaOoÃãÕõ",
    },
    "Achado Crítico": {
        "psm": 7,
        "scale": 4,
        "whitelist": "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzÀÁÂÃÉÊÍÓÔÕÚÇàáâãéêíóôõúç0123456789/:,.-() ",
    },
    "Data e Hora": {
        "psm": 7,
        "scale": 6,
        "whitelist": "0123456789/:- ",
    },
    "Observações": {
        "psm": 6,
        "scale": 4,
        "whitelist": None,
    },
}

class DashboardAchadosCriticos:
    def __init__(self):
        self.df_achados = None
        self.df_status = None
        self.df_correlacionado = None
        self.ris_ocr_backend = None
        self._rapidocr_engine = None

    def _parse_datetime_series(self, series):
        """Converte série para datetime priorizando padrão brasileiro."""
        parsed = pd.to_datetime(series, format='%d/%m/%Y %H:%M', errors='coerce')
        parsed = parsed.fillna(pd.to_datetime(series, format='%d-%m-%Y %H:%M', errors='coerce'))
        parsed = parsed.fillna(pd.to_datetime(series, format='%d/%m/%Y', errors='coerce'))
        parsed = parsed.fillna(pd.to_datetime(series, format='%d-%m-%Y', errors='coerce'))
        parsed = parsed.fillna(pd.to_datetime(series, errors='coerce', dayfirst=True))
        return parsed

    def _parse_datetime_value(self, value):
        """Converte valor único para datetime priorizando padrão brasileiro."""
        return self._parse_datetime_series(pd.Series([value])).iloc[0]

    def _normalize_text(self, value):
        """Normaliza texto para comparacao robusta entre planilhas."""
        if pd.isna(value):
            return ""
        text = str(value).strip().lower()
        text = "".join(
            char for char in unicodedata.normalize("NFKD", text)
            if not unicodedata.combining(char)
        )
        text = re.sub(r"[^a-z0-9]+", " ", text)
        return re.sub(r"\s+", " ", text).strip()

    def _normalize_identifier(self, value):
        """Normaliza identificadores como SAME preservando comparacao numerica."""
        if pd.isna(value):
            return ""
        if isinstance(value, (int, np.integer)):
            return str(int(value))
        if isinstance(value, (float, np.floating)) and np.isfinite(value):
            if float(value).is_integer():
                return str(int(value))

        text = str(value).strip()
        text = re.sub(r"\.0$", "", text)
        digits = re.sub(r"\D+", "", text)
        if digits:
            return digits.lstrip("0") or "0"
        return self._normalize_text(text)

    def _text_similarity(self, left, right):
        """Calcula similaridade textual normalizada."""
        left_norm = self._normalize_text(left)
        right_norm = self._normalize_text(right)
        if not left_norm or not right_norm:
            return 0.0
        if left_norm == right_norm:
            return 1.0
        return SequenceMatcher(None, left_norm, right_norm).ratio()

    def _token_overlap(self, left, right):
        """Mede sobreposicao de tokens, util para nomes de pacientes."""
        left_tokens = set(self._normalize_text(left).split())
        right_tokens = set(self._normalize_text(right).split())
        if not left_tokens or not right_tokens:
            return 0.0
        return len(left_tokens & right_tokens) / max(len(left_tokens | right_tokens), 1)

    def _format_date_columns(self, df):
        """Formata colunas de data/hora apenas para exibição/export."""
        df_fmt = df.copy()
        candidate_cols = {
            self.data_col_achados if hasattr(self, 'data_col_achados') else None,
            self.data_col_status if hasattr(self, 'data_col_status') else None,
            self.data_sinalizacao_col if hasattr(self, 'data_sinalizacao_col') else None,
            self.status_col if hasattr(self, 'status_col') else None,
            'Data_Exame',
            'Data_Sinalização',
            'DATA_HORA_PRESCRICAO',
            'STATUS_ALAUDAR',
            'data_sinalizacao_dt',
            'status_laudar_dt',
        }

        for col in candidate_cols:
            if col and col in df_fmt.columns:
                parsed = self._parse_datetime_series(df_fmt[col])
                has_time = (
                    parsed.dt.hour.fillna(0).ne(0) |
                    parsed.dt.minute.fillna(0).ne(0) |
                    parsed.dt.second.fillna(0).ne(0)
                ).any()
                date_format = '%d/%m/%Y %H:%M' if has_time else '%d/%m/%Y'
                formatted = parsed.dt.strftime(date_format)
                df_fmt[col] = formatted.where(parsed.notna(), '')

        return df_fmt

    def _ocr_dependencies_ready(self):
        """Verifica se as dependencias de OCR estao disponiveis."""
        self.ris_ocr_backend = None
        self._rapidocr_engine = None

        tesseract_error = None
        try:
            import pytesseract
            from PIL import Image  # noqa: F401
            tesseract_cmd = ADMIN_SMTP_CONFIG.get("tesseract_cmd")
            candidate_paths = [
                tesseract_cmd,
                shutil.which("tesseract"),
                "/usr/bin/tesseract",
                "/usr/local/bin/tesseract",
                "/opt/render/project/src/.apt/usr/bin/tesseract",
                "/opt/render/project/.apt/usr/bin/tesseract",
            ]
            executable = next((path for path in candidate_paths if path and os.path.exists(path)), None)
            if executable:
                pytesseract.pytesseract.tesseract_cmd = executable
                self.ris_ocr_backend = "tesseract"
                return True, ""
            tesseract_error = (
                "Executavel do Tesseract nao encontrado. "
                "Configure ADMIN_SMTP_CONFIG['tesseract_cmd'] ou instale o Tesseract."
            )
        except Exception as exc:
            tesseract_error = f"Dependencia Tesseract indisponivel: {exc}"

        try:
            from rapidocr_onnxruntime import RapidOCR

            self._rapidocr_engine = RapidOCR()
            self.ris_ocr_backend = "rapidocr"
            return True, ""
        except Exception as exc:
            rapidocr_error = f"Dependencia RapidOCR indisponivel: {exc}"

        return False, f"{tesseract_error} {rapidocr_error}".strip()

    def _detect_ris_dialog_bounds(self, image):
        """Localiza automaticamente a janela 'Resultado Crítico' no screenshot."""
        try:
            import cv2
        except Exception:
            return None

        rgb = np.array(image.convert("RGB"))
        hsv = cv2.cvtColor(rgb, cv2.COLOR_RGB2HSV)

        # Dialogo: baixo contraste de saturacao e brilho mais alto que o fundo.
        mask = cv2.inRange(hsv, np.array([0, 0, 95]), np.array([180, 90, 235]))
        kernel = np.ones((7, 7), np.uint8)
        mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
        mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)

        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        image_h, image_w = rgb.shape[:2]
        best_rect = None
        best_score = -1

        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            area = w * h
            if area < image_w * image_h * 0.01:
                continue
            if area > image_w * image_h * 0.35:
                continue
            if w < image_w * 0.15 or h < image_h * 0.10:
                continue
            if y > image_h * 0.75:
                continue
            if x < image_w * 0.02 or y < image_h * 0.02:
                continue
            if (x + w) > image_w * 0.98 or (y + h) > image_h * 0.98:
                continue

            aspect_ratio = w / max(h, 1)
            if not 1.1 <= aspect_ratio <= 3.8:
                continue

            contour_area = cv2.contourArea(contour)
            rectangularity = contour_area / max(area, 1)
            if rectangularity < 0.55:
                continue

            center_penalty = abs((x + w / 2) - image_w / 2) / image_w
            top_bonus = max(0, (image_h * 0.55 - y)) / image_h
            left_bias_bonus = max(0, (image_w * 0.7 - x)) / image_w
            score = (
                area
                + (rectangularity * area * 0.4)
                + (top_bonus * area * 0.25)
                + (left_bias_bonus * area * 0.15)
                - (center_penalty * area * 0.25)
            )
            if score > best_score:
                best_score = score
                best_rect = (x, y, x + w, y + h)

        return best_rect

    def _fallback_ris_dialog_bounds(self, image_width, image_height):
        """Fallback geométrico para screenshots do RIS com grande borda externa."""
        return (
            int(image_width * 0.05),
            int(image_height * 0.14),
            int(image_width * 0.88),
            int(image_height * 0.91),
        )

    def _relative_box_to_pixels(self, reference_box, relative_box, image_width, image_height):
        """Converte caixa relativa ao dialogo para pixels absolutos."""
        ref_x1, ref_y1, ref_x2, ref_y2 = reference_box
        ref_w = ref_x2 - ref_x1
        ref_h = ref_y2 - ref_y1
        rel_x1, rel_y1, rel_x2, rel_y2 = relative_box

        x1 = max(0, int(ref_x1 + (rel_x1 * ref_w)))
        y1 = max(0, int(ref_y1 + (rel_y1 * ref_h)))
        x2 = min(image_width, int(ref_x1 + (rel_x2 * ref_w)))
        y2 = min(image_height, int(ref_y1 + (rel_y2 * ref_h)))
        return (x1, y1, x2, y2)

    def _build_ris_debug_payload(self, image, dialog_box, extracted_items):
        """Monta artefatos visuais para calibracao do OCR."""
        annotated = image.convert("RGB").copy()
        draw = ImageDraw.Draw(annotated)
        draw.rectangle(dialog_box, outline=(0, 255, 255), width=3)

        debug_fields = []
        colors = [
            (255, 0, 0),
            (0, 255, 0),
            (255, 255, 0),
            (255, 128, 0),
            (255, 0, 255),
            (0, 200, 255),
        ]

        for idx, item in enumerate(extracted_items):
            box = item["box"]
            color = colors[idx % len(colors)]
            draw.rectangle(box, outline=color, width=3)
            draw.text((box[0] + 4, max(0, box[1] - 18)), item["field"], fill=color)
            debug_fields.append({
                "field": item["field"],
                "box": box,
                "raw_text": item["raw_text"],
                "value": item["value"],
                "crop": image.crop(box).copy(),
            })

        return {
            "annotated_image": annotated,
            "dialog_box": dialog_box,
            "fields": debug_fields,
        }

    def _clean_ocr_text(self, text):
        """Normaliza o texto extraido por OCR."""
        cleaned = " ".join(str(text).replace("\n", " ").split())
        return cleaned.replace("|", "").strip(" :-")

    def _normalize_ris_datetime(self, text):
        """Normaliza data/hora para padrão DD/MM/AAAA HH:MM quando possível."""
        cleaned = self._clean_ocr_text(text)
        cleaned = cleaned.replace("\\", "/").replace(".", ":")
        cleaned = re.sub(r"[^0-9/: -]", "", cleaned)
        match = re.search(r"(\d{2}/\d{2}/\d{2,4})(?:\s+(\d{2}:\d{2}))?", cleaned)
        if not match:
            return cleaned

        date_part = match.group(1)
        time_part = match.group(2)
        try:
            if len(date_part.split("/")[-1]) == 2:
                dt = datetime.strptime(date_part, "%d/%m/%y")
            else:
                dt = datetime.strptime(date_part, "%d/%m/%Y")
            formatted = dt.strftime("%d/%m/%Y")
            return f"{formatted} {time_part}".strip() if time_part else formatted
        except ValueError:
            return cleaned

    def _score_ris_candidate(self, field_name, text):
        """Pontua o resultado OCR de acordo com o formato esperado do campo."""
        cleaned = self._clean_ocr_text(text)
        if not cleaned:
            return -100

        score = len(cleaned)

        if field_name == "Diagnóstico":
            if re.search(r"[A-Za-zÀ-ÿ]+,\s*[A-Za-zÀ-ÿ]+", cleaned):
                score += 40
            if ":" in cleaned:
                score += 10
            if re.search(r"\d", cleaned):
                score -= 20

        elif field_name == "Resultado Crítico":
            if re.fullmatch(r"[A-Za-zÀ-ÿ ]{3,40}", cleaned):
                score += 25
            if re.search(r"\d", cleaned):
                score -= 20

        elif field_name == "Achado Crítico":
            if len(cleaned) >= 12:
                score += 25
            if re.search(r"[A-Za-zÀ-ÿ]{4,}", cleaned):
                score += 20

        elif field_name == "Contato":
            if re.fullmatch(r"[A-Za-zÀ-ÿ ]{5,60}", cleaned):
                score += 30
            if len(cleaned.split()) >= 2:
                score += 10
            if re.search(r"\d", cleaned):
                score -= 25

        elif field_name == "Contato com (Sucesso)":
            lowered = cleaned.lower()
            if lowered in {"sim", "não", "nao"}:
                score += 50
            else:
                score -= 30

        elif field_name == "Data e Hora":
            if re.search(r"\d{2}/\d{2}/\d{2,4}", cleaned):
                score += 40
            if re.search(r"\d{2}:\d{2}", cleaned):
                score += 20

        elif field_name == "Observações":
            if len(cleaned) >= 15:
                score += 20

        return score

    def _prepare_ris_variants(self, crop, field_name):
        """Gera variantes da imagem para aumentar a precisão do OCR."""
        rules = RIS_FIELD_OCR_RULES.get(field_name, {})
        scale = rules.get("scale", 4)

        if field_name == "Contato com (Sucesso)":
            focus_width = max(1, int(crop.width * 0.45))
            left_margin = max(0, int(crop.width * 0.03))
            top_margin = max(0, int(crop.height * 0.12))
            bottom_margin = max(top_margin + 1, int(crop.height * 0.88))
            crop = crop.crop((left_margin, top_margin, focus_width, bottom_margin))

        base = crop.convert("L")
        base = base.filter(ImageFilter.SHARPEN)
        base = base.resize((max(1, base.width * scale), max(1, base.height * scale)))

        variants = []

        hard = base.copy().point(lambda p: 255 if p > 165 else 0)
        variants.append(hard)

        soft = base.copy().point(lambda p: 255 if p > 145 else 0)
        variants.append(soft)

        variants.append(base.copy())
        return variants

    def _post_process_ris_text(self, field_name, text):
        """Ajusta o texto extraido para o significado esperado na tabela."""
        cleaned = self._clean_ocr_text(text)

        if field_name == "Diagnóstico":
            cleaned = re.sub(r"^\s*diagn[oó]stico\s*[:\-]?\s*", "", cleaned, flags=re.IGNORECASE)
            cleaned = re.sub(r"\s+[xX]\s*$", "", cleaned).strip()
            match = re.search(r"([A-Za-zÀ-ÿ]+,\s*[A-Za-zÀ-ÿ ]+)", cleaned)
            if match:
                cleaned = match.group(1).strip()
            elif len(cleaned) < 5 or len(re.findall(r"[A-Za-zÀ-ÿ]", cleaned)) < 4:
                return ""

        if field_name == "Resultado Crítico":
            cleaned = re.sub(r"^\s*(resultado\s*critico|comunicado)\s*[:\-]?\s*", "", cleaned, flags=re.IGNORECASE)

        if field_name == "Contato com (Sucesso)":
            lowered = cleaned.lower()
            if "sim" in lowered:
                return "Sim"
            if "nao" in lowered or "não" in lowered:
                return "Não"

        if field_name == "Achado Crítico":
            cleaned = re.sub(r"^\s*achado\s*critico\s*[:\-]?\s*", "", cleaned, flags=re.IGNORECASE)

        if field_name == "Data e Hora":
            return self._normalize_ris_datetime(cleaned)

        return cleaned

    def _ocr_ris_region(self, image, box, field_name, multiline=False):
        """Executa OCR em uma regiao da tela RIS."""
        crop = image.crop(box)
        rules = RIS_FIELD_OCR_RULES.get(field_name, {})
        variants = self._prepare_ris_variants(crop, field_name)

        if self.ris_ocr_backend == "tesseract":
            import pytesseract

            best_text = ""
            best_score = -1000
            psm = str(rules.get("psm", 6 if multiline else 7))
            whitelist = rules.get("whitelist")
            config_base = f"--oem 3 --psm {psm} -c preserve_interword_spaces=1"
            if whitelist:
                config_base += f' -c tessedit_char_whitelist="{whitelist}"'

            for variant in variants:
                text = pytesseract.image_to_string(
                    variant,
                    lang="por+eng",
                    config=config_base
                )
                score = self._score_ris_candidate(field_name, text)
                if score > best_score:
                    best_score = score
                    best_text = text

            if field_name == "Contato com (Sucesso)" and not self._clean_ocr_text(best_text):
                fallback_config = "--oem 3 --psm 8"
                for variant in variants:
                    text = pytesseract.image_to_string(
                        variant,
                        lang="por+eng",
                        config=fallback_config
                    )
                    score = self._score_ris_candidate(field_name, text)
                    if score > best_score:
                        best_score = score
                        best_text = text

            return self._clean_ocr_text(best_text)

        if self.ris_ocr_backend == "rapidocr":
            best_text = ""
            best_score = -1000

            for variant in variants:
                image_array = np.array(variant.convert("RGB"))
                result = self._rapidocr_engine(
                    image_array,
                    use_det=False,
                    use_cls=False,
                    use_rec=True
                )
                texts = getattr(result, "txts", None)
                if texts is None and isinstance(result, tuple) and result:
                    texts = result[0]
                if isinstance(texts, str):
                    text = texts
                elif texts:
                    text = " ".join(str(item) for item in texts if item)
                else:
                    text = ""

                score = self._score_ris_candidate(field_name, text)
                if score > best_score:
                    best_score = score
                    best_text = text

            return self._clean_ocr_text(best_text)

        return ""

    def extract_ris_screen_text(self, uploaded_image, include_debug=False):
        """Extrai campos preenchidos da tela RIS via OCR."""
        deps_ok, error_message = self._ocr_dependencies_ready()
        if not deps_ok:
            return None, error_message, None

        try:
            from PIL import Image

            image = Image.open(io.BytesIO(uploaded_image.getvalue()))
            width, height = image.size
            dialog_box = self._detect_ris_dialog_bounds(image)
            fallback_box = self._fallback_ris_dialog_bounds(width, height)

            if dialog_box:
                det_area = max(1, (dialog_box[2] - dialog_box[0]) * (dialog_box[3] - dialog_box[1]))
                fallback_area = max(1, (fallback_box[2] - fallback_box[0]) * (fallback_box[3] - fallback_box[1]))
                if det_area < fallback_area * 0.55:
                    dialog_box = fallback_box
            else:
                dialog_box = fallback_box

            extracted_rows = []
            debug_items = []

            diagnosis_box = self._relative_box_to_pixels(dialog_box, RIS_DIAGNOSIS_RELATIVE_BOX, width, height)
            diagnosis_text = self._ocr_ris_region(image, diagnosis_box, "Diagnóstico", multiline=False)
            diagnosis_value = self._post_process_ris_text("Diagnóstico", diagnosis_text)
            extracted_rows.append({
                "Campo": "Diagnóstico",
                "Valor": diagnosis_value
            })
            debug_items.append({
                "field": "Diagnóstico",
                "box": diagnosis_box,
                "raw_text": diagnosis_text,
                "value": diagnosis_value,
            })

            for region in RIS_DIALOG_FIELD_REGIONS:
                scaled_box = self._relative_box_to_pixels(dialog_box, region["box"], width, height)
                text = self._ocr_ris_region(
                    image,
                    scaled_box,
                    region["field"],
                    multiline=region.get("multiline", False)
                )
                value = self._post_process_ris_text(region["field"], text)
                extracted_rows.append({
                    "Campo": region["field"],
                    "Valor": value
                })
                debug_items.append({
                    "field": region["field"],
                    "box": scaled_box,
                    "raw_text": text,
                    "value": value,
                })

            debug_payload = self._build_ris_debug_payload(image, dialog_box, debug_items) if include_debug else None
            return pd.DataFrame(extracted_rows), None, debug_payload
        except Exception as exc:
            return None, f"Erro ao ler screenshot: {exc}", None

    def _smtp_config_ready(self):
        """Valida a configuracao de SMTP."""
        required_keys = ["host", "port", "sender_email"]
        missing = [key for key in required_keys if not ADMIN_SMTP_CONFIG.get(key)]
        if missing:
            return False, (
                "Configuracao SMTP incompleta. Defina as variaveis: "
                "SMTP_HOST, SMTP_PORT e SMTP_SENDER_EMAIL."
            )

        if ADMIN_SMTP_CONFIG.get("username") and not ADMIN_SMTP_CONFIG.get("password"):
            return False, "SMTP_USERNAME foi definido, mas SMTP_PASSWORD esta vazio."

        return True, ""

    def _create_ris_excel_attachment(self, table_df):
        """Gera anexo Excel com a tabela extraida."""
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            table_df.to_excel(writer, sheet_name='Leitura_RIS', index=False)
        output.seek(0)
        return output.getvalue()

    def send_ris_table_email(self, table_df, recipient):
        """Envia a tabela extraida para o email informado pelo usuario."""
        if not recipient:
            return False, "Informe o email de destino."

        if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", recipient):
            return False, "Informe um endereco de email valido."

        config_ok, config_error = self._smtp_config_ready()
        if not config_ok:
            return False, config_error

        if table_df is None or table_df.empty:
            return False, "Nenhum campo preenchido foi encontrado para envio."

        table_df = table_df.copy()
        table_df["Campo"] = table_df["Campo"].fillna("").astype(str).str.strip()
        table_df["Valor"] = table_df["Valor"].fillna("").astype(str).str.strip()
        table_df = table_df[(table_df["Campo"] != "") & (table_df["Valor"] != "")]

        if table_df.empty:
            return False, "Nenhum campo preenchido foi encontrado para envio."

        subject = f"Leitura da tela RIS - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        html_table = table_df.to_html(index=False, border=0)
        plain_table = table_df.to_string(index=False)
        excel_payload = self._create_ris_excel_attachment(table_df)

        message = EmailMessage()
        sender_name = ADMIN_SMTP_CONFIG.get("sender_name", "CDI - Achados Criticos")
        sender_email = ADMIN_SMTP_CONFIG["sender_email"]
        message["Subject"] = subject
        message["From"] = f"{sender_name} <{sender_email}>"
        message["To"] = recipient
        message.set_content(
            "Segue a tabela extraida da tela RIS.\n\n"
            f"{plain_table}\n\n"
            "Mensagem gerada automaticamente pelo dashboard de Achados Criticos."
        )
        message.add_alternative(
            f"""
            <html>
              <body>
                <p>Segue a tabela extraida da tela RIS.</p>
                {html_table}
                <p>Mensagem gerada automaticamente pelo dashboard de Achados Criticos.</p>
              </body>
            </html>
            """,
            subtype="html"
        )
        message.add_attachment(
            excel_payload,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=f"ris_tela_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        try:
            with smtplib.SMTP(ADMIN_SMTP_CONFIG["host"], ADMIN_SMTP_CONFIG["port"], timeout=30) as smtp:
                if ADMIN_SMTP_CONFIG.get("use_tls", True):
                    smtp.starttls()
                username = ADMIN_SMTP_CONFIG.get("username")
                password = ADMIN_SMTP_CONFIG.get("password")
                if username and password:
                    smtp.login(username, password)
                smtp.send_message(message)
            return True, f"Tabela enviada com sucesso para {recipient}."
        except Exception as exc:
            return False, f"Erro ao enviar email: {exc}"

    def render_ris_ocr_email_tab(self):
        """Renderiza aba de OCR da tela RIS e envio por email."""
        st.markdown("## Leitura da Tela RIS")
        st.caption("Carregue um screenshot da tela RIS, revise os campos extraidos e envie a tabela por email.")

        uploaded_image = st.file_uploader(
            "Screenshot da tela RIS",
            type=["png", "jpg", "jpeg", "bmp", "tif", "tiff"],
            key="ris_screenshot"
        )
        debug_mode = st.checkbox("Mostrar debug OCR", key="ris_debug_mode")

        if uploaded_image is not None:
            image_signature = f"{uploaded_image.name}:{uploaded_image.size}"
            if st.session_state.get("ris_image_signature") != image_signature:
                st.session_state.ris_image_signature = image_signature
                st.session_state.ris_extracted_df = pd.DataFrame(columns=["Campo", "Valor"])
                st.session_state.ris_debug_payload = None
                if "ris_editor" in st.session_state:
                    del st.session_state["ris_editor"]

            col_preview, col_actions = st.columns([1.35, 1])

            with col_preview:
                st.image(uploaded_image, caption="Preview do screenshot RIS", use_container_width=True)

            with col_actions:
                st.markdown("### Extracao")
                if st.button("Ler Tela RIS", type="primary", key="ris_extract_button"):
                    extracted_df, error_message, debug_payload = self.extract_ris_screen_text(
                        uploaded_image,
                        include_debug=debug_mode
                    )
                    if error_message:
                        st.error(error_message)
                    else:
                        st.session_state.ris_extracted_df = extracted_df
                        st.session_state.ris_debug_payload = debug_payload
                        backend_name = "Tesseract" if self.ris_ocr_backend == "tesseract" else "RapidOCR"
                        st.success(f"Leitura concluida. {len(extracted_df)} campo(s) foram enviados para revisao.")
                        st.caption(f"Backend OCR usado: {backend_name}")

                st.markdown("### Envio")
                recipient = st.text_input(
                    "Email destino",
                    key="ris_recipient_email",
                    placeholder="usuario@dominio.com"
                )

                current_df = st.session_state.get("ris_extracted_df", pd.DataFrame(columns=["Campo", "Valor"]))
                if st.button("Enviar Tabela por Email", key="ris_send_button"):
                    success, message = self.send_ris_table_email(current_df, recipient)
                    if success:
                        st.success(message)
                    else:
                        st.error(message)

            current_df = st.session_state.get("ris_extracted_df", pd.DataFrame(columns=["Campo", "Valor"]))
            if not current_df.empty:
                st.markdown("### Tabela Extraida")
                edited_df = st.data_editor(
                    current_df,
                    key="ris_editor",
                    use_container_width=True,
                    hide_index=True,
                    num_rows="fixed",
                    column_config={
                        "Campo": st.column_config.TextColumn("Campo", disabled=True),
                        "Valor": st.column_config.TextColumn("Valor"),
                    }
                )
                st.session_state.ris_extracted_df = pd.DataFrame(edited_df)
                st.download_button(
                    label="Baixar Tabela em Excel",
                    data=self._create_ris_excel_attachment(st.session_state.ris_extracted_df),
                    file_name=f"ris_tela_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="ris_download_button"
                )

                debug_payload = st.session_state.get("ris_debug_payload")
                if debug_mode and debug_payload:
                    st.markdown("### Debug OCR")
                    st.image(
                        debug_payload["annotated_image"],
                        caption="Caixa detectada e recortes usados no OCR",
                        use_container_width=True
                    )
                    for item in debug_payload["fields"]:
                        st.markdown(f"**{item['field']}**")
                        st.image(item["crop"], use_container_width=True)
                        st.caption(f"Bruto: {item['raw_text'] or '(vazio)'}")
                        st.caption(f"Final: {item['value'] or '(vazio)'}")
            else:
                st.info("Nenhum campo extraido ainda. Envie um screenshot e clique em 'Ler Tela RIS'.")
        else:
            st.info("Selecione um screenshot para habilitar a leitura OCR da tela RIS.")

    def render_header(self):
        """Renderiza o cabeçalho principal"""
        st.markdown("""
        <div class="main-header">
            <h1>🏥 CDI - Dashboard de Achados Críticos</h1>
            <p>Sistema Avançado de Monitoramento e Análise de Comunicação de Achados Críticos</p>
        </div>
        """, unsafe_allow_html=True)

    def render_sidebar(self):
        """Renderiza a sidebar com controles"""
        st.sidebar.markdown("## ⚙️ Configurações")

        # Upload de arquivos
        st.sidebar.markdown("### 📁 Upload de Planilhas")

        uploaded_achados = st.sidebar.file_uploader(
            "Planilha de Achados Críticos",
            type=['xlsx', 'xls'],
            key="achados",
            help="Selecione a planilha com os achados críticos reportados"
        )

        uploaded_status = st.sidebar.file_uploader(
            "Planilha de Status dos Exames",
            type=['xlsx', 'xls'],
            key="status",
            help="Selecione a planilha com os status dos exames"
        )

        return uploaded_achados, uploaded_status

    def render_date_filter(self):
        """Renderiza o filtro de data na sidebar"""
        if not st.session_state.processed or self.df_correlacionado is None:
            return None, None

        st.sidebar.markdown("---")
        st.sidebar.markdown("### 📅 Filtro de Período")

        # Extrair datas disponíveis
        if self.data_sinalizacao_col and self.data_sinalizacao_col in self.df_correlacionado.columns:
            datas = self._parse_datetime_series(self.df_correlacionado[self.data_sinalizacao_col])
            datas_validas = datas[datas.notna()]

            if len(datas_validas) > 0:
                # Obter anos e meses únicos
                anos_disponiveis = sorted(datas_validas.dt.year.unique(), reverse=True)

                # Seletor de ano
                ano_selecionado = st.sidebar.selectbox(
                    "Ano",
                    options=anos_disponiveis,
                    key="dashboard_filter_year",
                    help="Selecione o ano para análise"
                )

                # Filtrar meses do ano selecionado
                meses_do_ano = datas_validas[datas_validas.dt.year == ano_selecionado].dt.month.unique()
                meses_do_ano = sorted(meses_do_ano)

                # Mapear números para nomes de meses
                nomes_meses = {
                    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
                    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
                    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
                }

                opcoes_meses = ["Todos"] + [f"{nomes_meses[m]} ({m:02d})" for m in meses_do_ano]

                # Seletor de mês
                mes_selecionado = st.sidebar.selectbox(
                    "Mês",
                    options=opcoes_meses,
                    key="dashboard_filter_month",
                    help="Selecione o mês para análise ou 'Todos' para o ano inteiro"
                )

                # Extrair número do mês
                if mes_selecionado == "Todos":
                    mes_numero = None
                else:
                    mes_numero = int(mes_selecionado.split("(")[1].split(")")[0])

                # Mostrar informação do filtro
                if mes_numero:
                    st.sidebar.info(f"📊 Analisando: **{nomes_meses[mes_numero]} {ano_selecionado}**")
                else:
                    st.sidebar.info(f"📊 Analisando: **Todo o ano {ano_selecionado}**")

                return ano_selecionado, mes_numero

        return None, None

    def apply_date_filter(self, ano_filtro, mes_filtro):
        """Aplica filtro de data sempre sobre o dataframe completo da sessão."""
        if self.df_correlacionado is None or ano_filtro is None or not self.data_sinalizacao_col:
            return 0, 0

        base_df = self.df_correlacionado.copy()
        datas = self._parse_datetime_series(base_df[self.data_sinalizacao_col])
        mask_final = datas.dt.year == ano_filtro

        if mes_filtro is not None:
            mask_final = mask_final & (datas.dt.month == mes_filtro)

        self.df_correlacionado = base_df[mask_final].copy()

        review_df = getattr(self, 'df_revisao_correlacao', None)
        if review_df is not None and not review_df.empty and self.data_sinalizacao_col in review_df.columns:
            review_datas = self._parse_datetime_series(review_df[self.data_sinalizacao_col])
            review_mask = review_datas.dt.year == ano_filtro
            if mes_filtro is not None:
                review_mask = review_mask & (review_datas.dt.month == mes_filtro)
            self.df_revisao_correlacao = review_df[review_mask].copy()

        return len(self.df_correlacionado), len(base_df)

    def load_data(self, uploaded_achados, uploaded_status):
        """Carrega e processa os dados das planilhas"""
        try:
            if uploaded_achados is not None:
                # Detectar engine baseado na extensão
                engine = 'openpyxl' if uploaded_achados.name.endswith('.xlsx') else 'xlrd'
                self.df_achados = pd.read_excel(uploaded_achados, engine=engine)
                self.df_achados = self.df_achados.dropna(how='all')

            if uploaded_status is not None:
                # Carregar planilha de status com header padrão
                engine = 'openpyxl' if uploaded_status.name.endswith('.xlsx') else 'xlrd'
                self.df_status = pd.read_excel(uploaded_status, engine=engine)
                self.df_status = self.df_status.dropna(how='all')

            return True
        except Exception as e:
            st.error(f"❌ Erro ao carregar planilhas: {str(e)}")
            import traceback
            st.error(traceback.format_exc())
            return False

    def identify_columns(self):
        """Identifica automaticamente as colunas necessárias"""
        # Encontrar coluna SAME
        self.same_col_achados = None
        self.same_col_status = None
        for col in self.df_achados.columns:
            if isinstance(col, str) and 'same' in self._normalize_text(col):
                self.same_col_achados = col
                break

        for col in self.df_status.columns:
            if isinstance(col, str) and 'same' in self._normalize_text(col):
                self.same_col_status = col
                break

        # Encontrar coluna Nome do Paciente
        self.nome_col_achados = None
        self.nome_col_status = None
        for col in self.df_achados.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'nome' in col_norm and 'paciente' in col_norm:
                self.nome_col_achados = col
                break

        for col in self.df_status.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'nome' in col_norm and 'paciente' in col_norm:
                self.nome_col_status = col
                break

        # Encontrar coluna Data Exame
        self.data_col_achados = None
        self.data_col_status = None
        for col in self.df_achados.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'data' in col_norm and 'exame' in col_norm:
                self.data_col_achados = col
                break

        for col in self.df_status.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'data' in col_norm and ('prescri' in col_norm or 'hora' in col_norm):
                self.data_col_status = col
                break

        # Encontrar coluna Descrição Procedimento
        self.desc_col_achados = None
        self.desc_col_status = None
        for col in self.df_achados.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'descri' in col_norm and 'proced' in col_norm:
                self.desc_col_achados = col
                break

        for col in self.df_status.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'descri' in col_norm and 'proced' in col_norm:
                self.desc_col_status = col
                break

        # Encontrar coluna Modalidade
        self.modalidade_col_achados = None
        self.modalidade_col_status = None
        for col in self.df_achados.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'modalidade' in col_norm:
                self.modalidade_col_achados = col
                break

        for col in self.df_status.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'modalidade' in col_norm:
                self.modalidade_col_status = col
                break

        # Encontrar STATUS_ALAUDAR
        self.status_col = None
        for col in self.df_status.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'status' in col_norm and 'laudar' in col_norm:
                self.status_col = col
                break

        # Encontrar Data Sinalização
        self.data_sinalizacao_col = None
        for col in self.df_achados.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'data' in col_norm and 'sinalizacao' in col_norm:
                self.data_sinalizacao_col = col
                break

        # Encontrar coluna Medico Laudo
        self.medico_col = None
        for col in self.df_achados.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'medico' in col_norm and 'laudo' in col_norm:
                self.medico_col = col
                break

        # Encontrar coluna Achado Crítico
        self.achado_col = None
        for col in self.df_achados.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'achado' in col_norm and 'critico' in col_norm:
                self.achado_col = col
                break

        # Encontrar coluna Contato (médico que recebeu o contato)
        self.contato_col = None
        for col in self.df_achados.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'contato' in col_norm and 'sucesso' not in col_norm:
                self.contato_col = col
                break

        # Encontrar coluna Informado Por (médico que fez a comunicação)
        self.informado_por_col = None
        for col in self.df_achados.columns:
            col_norm = self._normalize_text(col)
            if isinstance(col, str) and 'informado' in col_norm and 'por' in col_norm:
                self.informado_por_col = col
                break

        return (self.same_col_achados is not None and
                self.same_col_status is not None)

    def _score_status_candidate(self, achado, exame):
        """Pontua um exame candidato para pareamento com um achado critico."""
        score = 100.0

        achado_nome = achado.get(self.nome_col_achados, "") if self.nome_col_achados else ""
        exame_nome = exame.get(self.nome_col_status, "") if self.nome_col_status else ""
        name_similarity = max(
            self._text_similarity(achado_nome, exame_nome),
            self._token_overlap(achado_nome, exame_nome),
        )
        score += name_similarity * 30

        achado_proc = achado.get(self.desc_col_achados, "") if self.desc_col_achados else ""
        exame_proc = exame.get(self.desc_col_status, "") if self.desc_col_status else ""
        procedure_similarity = self._text_similarity(achado_proc, exame_proc)
        if self.desc_col_achados and self.desc_col_status:
            score += procedure_similarity * 90
            if procedure_similarity >= 0.95:
                score += 20
            elif procedure_similarity < 0.45:
                score -= 25

        achado_dt = self._parse_datetime_value(
            achado.get(self.data_col_achados)
        ) if self.data_col_achados else pd.NaT
        exame_dt = self._parse_datetime_value(
            exame.get(self.data_col_status)
        ) if self.data_col_status else pd.NaT

        same_day = False
        time_delta_min = np.nan
        if pd.notna(achado_dt) and pd.notna(exame_dt):
            same_day = achado_dt.date() == exame_dt.date()
            time_delta_min = abs((exame_dt - achado_dt).total_seconds()) / 60
            if same_day:
                score += 60
            if time_delta_min <= 1:
                score += 30
            elif time_delta_min <= 60:
                score += 22
            elif time_delta_min <= 360:
                score += 14
            elif time_delta_min <= 1440:
                score += 6
            else:
                score -= min(35, time_delta_min / 1440)

        modalidade_similarity = 0.0
        if self.modalidade_col_achados and self.modalidade_col_status:
            modalidade_similarity = self._text_similarity(
                achado.get(self.modalidade_col_achados, ""),
                exame.get(self.modalidade_col_status, ""),
            )
            score += modalidade_similarity * 10

        if self.status_col and pd.notna(exame.get(self.status_col)):
            score += 5

        return {
            "score": round(float(score), 4),
            "name_similarity": round(float(name_similarity), 4),
            "procedure_similarity": round(float(procedure_similarity), 4),
            "modality_similarity": round(float(modalidade_similarity), 4),
            "same_day": bool(same_day),
            "time_delta_min": round(float(time_delta_min), 2) if pd.notna(time_delta_min) else np.nan,
        }

    def _match_is_reliable(self, details):
        """Decide se o melhor candidato pode ser usado no calculo operacional."""
        has_procedure_cols = bool(self.desc_col_achados and self.desc_col_status)
        if has_procedure_cols and details["procedure_similarity"] < 0.72:
            return False

        has_name_cols = bool(self.nome_col_achados and self.nome_col_status)
        if has_name_cols and details["name_similarity"] < 0.35:
            return False

        time_delta = details["time_delta_min"]
        if pd.notna(time_delta) and time_delta > 1440 and not details["same_day"]:
            return False

        return details["score"] >= 160

    def _match_confidence(self, details):
        """Classifica a confianca do pareamento."""
        if not self._match_is_reliable(details):
            return "Baixa"
        if (
            details["score"] >= 245 and
            details["procedure_similarity"] >= 0.90 and
            details["same_day"]
        ):
            return "Alta"
        return "Media"

    def correlate_data(self):
        """Correlaciona as planilhas usando múltiplos critérios"""
        if self.df_achados is None or self.df_status is None:
            return False

        try:
            # Identificar colunas
            if not self.identify_columns():
                st.error("❌ Não foi possível identificar as colunas necessárias")
                return False

            status_work = self.df_status.copy()
            status_work["_match_same_norm"] = status_work[self.same_col_status].map(self._normalize_identifier)

            correlacoes = []
            match_counts = {
                "alta": 0,
                "media": 0,
                "baixa": 0,
                "sem_same": 0,
                "sem_confianca": 0,
            }

            for _, achado in self.df_achados.iterrows():
                same_achado_norm = self._normalize_identifier(achado[self.same_col_achados])

                # Buscar exames correspondentes com SAME normalizado.
                exames_same = status_work[status_work["_match_same_norm"] == same_achado_norm]

                correlacao = achado.copy()
                correlacao["match_status"] = ""
                correlacao["match_confidence"] = ""
                correlacao["match_score"] = np.nan
                correlacao["match_name_similarity"] = np.nan
                correlacao["match_procedure_similarity"] = np.nan
                correlacao["match_modality_similarity"] = np.nan
                correlacao["match_same_day"] = False
                correlacao["match_time_delta_min"] = np.nan
                correlacao["match_candidates"] = len(exames_same)
                correlacao["match_status_row"] = np.nan
                correlacao["match_candidate_procedure"] = ""
                correlacao["match_candidate_datetime"] = ""

                if len(exames_same) == 0:
                    correlacao["match_status"] = "sem_same"
                    match_counts["sem_same"] += 1
                    correlacoes.append(correlacao)
                    continue

                scored_candidates = []
                for status_idx, exame in exames_same.iterrows():
                    details = self._score_status_candidate(achado, exame)
                    scored_candidates.append((details["score"], status_idx, exame, details))

                scored_candidates.sort(key=lambda item: item[0], reverse=True)
                _, status_idx, melhor_exame, details = scored_candidates[0]

                reliable = self._match_is_reliable(details)
                confidence = self._match_confidence(details)
                correlacao["match_confidence"] = confidence
                correlacao["match_score"] = details["score"]
                correlacao["match_name_similarity"] = details["name_similarity"]
                correlacao["match_procedure_similarity"] = details["procedure_similarity"]
                correlacao["match_modality_similarity"] = details["modality_similarity"]
                correlacao["match_same_day"] = details["same_day"]
                correlacao["match_time_delta_min"] = details["time_delta_min"]
                correlacao["match_status_row"] = status_idx
                correlacao["match_candidate_procedure"] = (
                    melhor_exame.get(self.desc_col_status, "") if self.desc_col_status else ""
                )
                correlacao["match_candidate_datetime"] = (
                    melhor_exame.get(self.data_col_status, "") if self.data_col_status else ""
                )

                if not reliable:
                    correlacao["match_status"] = "sem_correlacao_confiavel"
                    match_counts["baixa"] += 1
                    match_counts["sem_confianca"] += 1
                    correlacoes.append(correlacao)
                    continue

                correlacao["match_status"] = "correlacionado"
                if confidence == "Alta":
                    match_counts["alta"] += 1
                else:
                    match_counts["media"] += 1

                for col in melhor_exame.index:
                    if col in {"_match_same_norm"}:
                        continue
                    if col != self.same_col_status:
                        correlacao[col] = melhor_exame[col]

                correlacoes.append(correlacao)

            self.df_correlacionado = pd.DataFrame(correlacoes)
            self.match_summary = match_counts

            total_ok = match_counts["alta"] + match_counts["media"]
            st.sidebar.info(
                f"🔗 Correlação confiável: {total_ok}/{len(self.df_achados)} "
                f"registro(s). Revisar: {match_counts['sem_same'] + match_counts['sem_confianca']}."
            )
            return True

        except Exception as e:
            st.error(f"❌ Erro na correlação: {str(e)}")
            return False

    def calculate_times(self):
        """Calcula os tempos de comunicação"""
        if self.df_correlacionado is None:
            return False

        try:
            # Verificar se as colunas necessárias existem
            if not self.status_col or not self.data_sinalizacao_col:
                st.error("❌ Colunas necessárias não encontradas para cálculo de tempos")
                return False

            # Filtrar registros com STATUS_ALAUDAR válido
            df_com_status = self.df_correlacionado[
                self.df_correlacionado[self.status_col].notna()
            ].copy()

            registros_sem_status = len(self.df_correlacionado) - len(df_com_status)
            self.df_revisao_correlacao = self.df_correlacionado[
                self.df_correlacionado[self.status_col].isna()
            ].copy()
            self.time_summary = {
                "total_correlacionado": len(self.df_correlacionado),
                "sem_status": registros_sem_status,
                "datas_invalidas": 0,
                "calculados": 0,
            }
            if registros_sem_status > 0:
                st.sidebar.warning(
                    f"⚠️ {registros_sem_status} registro(s) sem correlação/status confiável "
                    "não entraram no cálculo de tempo"
                )

            if len(df_com_status) == 0:
                st.warning("⚠️ Nenhum registro com STATUS_ALAUDAR encontrado")
                return False

            # Converter datas
            df_com_status['data_sinalizacao_dt'] = self._parse_datetime_series(
                df_com_status[self.data_sinalizacao_col]
            )
            df_com_status['status_laudar_dt'] = self._parse_datetime_series(
                df_com_status[self.status_col]
            )

            # Calcular tempos
            df_com_status['tempo_comunicacao_horas'] = (
                df_com_status['data_sinalizacao_dt'] - df_com_status['status_laudar_dt']
            ).dt.total_seconds() / 3600

            df_com_status['fora_do_prazo'] = df_com_status['tempo_comunicacao_horas'] > 1

            # Filtrar registros válidos
            # NOTA: Removemos filtro de tempo >= 0 para maior sensibilidade
            # Tempos negativos podem indicar comunicação antes de laudo finalizado (casos especiais)
            registros_validos = (
                df_com_status['data_sinalizacao_dt'].notna() &
                df_com_status['status_laudar_dt'].notna()
            )

            datas_invalidas = len(df_com_status) - int(registros_validos.sum())
            self.time_summary["datas_invalidas"] = datas_invalidas
            if datas_invalidas > 0:
                self.df_revisao_correlacao = pd.concat(
                    [
                        self.df_revisao_correlacao,
                        df_com_status[~registros_validos].copy(),
                    ],
                    ignore_index=True,
                    sort=False,
                )
            if datas_invalidas > 0:
                st.sidebar.warning(
                    f"⚠️ {datas_invalidas} registro(s) com datas inválidas removido(s)"
                )

            self.df_correlacionado = df_com_status[registros_validos].copy()
            self.time_summary["calculados"] = len(self.df_correlacionado)

            # Criar coluna adicional para indicar casos com tempo negativo
            self.df_correlacionado['tempo_negativo'] = self.df_correlacionado['tempo_comunicacao_horas'] < 0

            # Remover registros sem médico
            self.remove_without_doctor()

            # Remover duplicidades
            self.remove_duplicates()

            return True

        except Exception as e:
            st.error(f"❌ Erro no cálculo de tempos: {str(e)}")
            return False

    def remove_without_doctor(self):
        """Remove registros que não possuem nome do médico"""
        if self.df_correlacionado is None or len(self.df_correlacionado) == 0:
            return

        if not self.medico_col:
            return

        total_antes = len(self.df_correlacionado)

        # Filtrar apenas registros com médico preenchido
        self.df_correlacionado = self.df_correlacionado[
            self.df_correlacionado[self.medico_col].notna() &
            (self.df_correlacionado[self.medico_col].astype(str).str.strip() != '')
        ].copy()

        total_depois = len(self.df_correlacionado)
        removidos = total_antes - total_depois

        if removidos > 0:
            st.sidebar.warning(f"⚠️ {removidos} registro(s) sem médico removido(s)")

    def remove_duplicates(self):
        """Remove registros duplicados COMPLETOS (incluindo mesmo médico) mantendo o de menor tempo"""
        if self.df_correlacionado is None or len(self.df_correlacionado) == 0:
            return

        total_antes = len(self.df_correlacionado)

        # Identificar colunas para detectar duplicatas COMPLETAS
        cols_duplicata = []
        if self.same_col_achados:
            cols_duplicata.append(self.same_col_achados)
        if self.data_sinalizacao_col:
            cols_duplicata.append(self.data_sinalizacao_col)
        if self.achado_col:
            cols_duplicata.append(self.achado_col)

        # IMPORTANTE: Adicionar médico como critério de duplicata
        # Se médico for diferente, NÃO é duplicata
        if self.medico_col:
            cols_duplicata.append(self.medico_col)

        if not cols_duplicata:
            return

        # Ordenar por colunas de duplicata + tempo (menor primeiro)
        sort_cols = cols_duplicata + ['tempo_comunicacao_horas']
        sort_ascending = [True] * len(cols_duplicata) + [True]  # Tempo crescente (menor primeiro)

        self.df_correlacionado = self.df_correlacionado.sort_values(
            by=sort_cols,
            ascending=sort_ascending
        )

        # Manter apenas a primeira ocorrência (menor tempo)
        # Só remove se TUDO for igual (SAME + Data + Achado + MÉDICO)
        self.df_correlacionado = self.df_correlacionado.drop_duplicates(
            subset=cols_duplicata,
            keep='first'
        )

        total_depois = len(self.df_correlacionado)
        duplicatas_removidas = total_antes - total_depois

        if duplicatas_removidas > 0:
            st.sidebar.info(f"🔄 {duplicatas_removidas} duplicata(s) removida(s)")

        return duplicatas_removidas

    def render_metrics_overview(self):
        """Renderiza as métricas principais"""
        if self.df_correlacionado is None or len(self.df_correlacionado) == 0:
            return

        total_casos = len(self.df_correlacionado)
        no_prazo = (~self.df_correlacionado['fora_do_prazo']).sum()
        fora_prazo = self.df_correlacionado['fora_do_prazo'].sum()
        percentual_compliance = (no_prazo / total_casos) * 100

        tempo_mediano = self.df_correlacionado['tempo_comunicacao_horas'].median()
        tempo_medio = self.df_correlacionado['tempo_comunicacao_horas'].mean()

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            compliance_color = "success" if percentual_compliance >= 80 else "warning" if percentual_compliance >= 60 else "danger"
            st.markdown(f"""
            <div class="big-metric">
                <h2 class="status-{compliance_color}">{percentual_compliance:.1f}%</h2>
                <p>Taxa de Compliance</p>
            </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
            <div class="big-metric">
                <h2 class="status-success">{no_prazo}</h2>
                <p>No Prazo (≤ 1h)</p>
            </div>
            """, unsafe_allow_html=True)

        with col3:
            st.markdown(f"""
            <div class="big-metric">
                <h2 class="status-danger">{fora_prazo}</h2>
                <p>Fora do Prazo (> 1h)</p>
            </div>
            """, unsafe_allow_html=True)

        with col4:
            tempo_color = "success" if tempo_mediano <= 0.5 else "warning" if tempo_mediano <= 1 else "danger"
            st.markdown(f"""
            <div class="big-metric">
                <h2 class="status-{tempo_color}">{tempo_mediano:.1f}h</h2>
                <p>Tempo Mediano</p>
            </div>
            """, unsafe_allow_html=True)

    def create_compliance_chart(self):
        """Cria gráfico de compliance"""
        if self.df_correlacionado is None:
            return None

        no_prazo = (~self.df_correlacionado['fora_do_prazo']).sum()
        fora_prazo = self.df_correlacionado['fora_do_prazo'].sum()
        total = no_prazo + fora_prazo
        compliance_pct = (no_prazo / total * 100) if total else 0

        fig = go.Figure(data=[
            go.Pie(
                labels=['No Prazo (≤ 1h)', 'Fora do Prazo (> 1h)'],
                values=[no_prazo, fora_prazo],
                hole=0.6,
                marker=dict(
                    colors=['#27AE60', '#E74C3C'],
                    line=dict(color='#2C3E50', width=2)
                ),
                textinfo='label+percent+value',
                textfont=dict(size=14, color='white'),
                hovertemplate='<b>%{label}</b><br>Casos: %{value}<br>Percentual: %{percent}<extra></extra>'
            )
        ])

        fig.update_layout(
            title=dict(
                text='Distribuição de Compliance',
                x=0.5,
                font=dict(size=20, color='white')
            ),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.2,
                xanchor="center",
                x=0.5,
                font=dict(color='white')
            ),
            margin=dict(t=80, b=80, l=40, r=40),
            annotations=[
                dict(
                    text=f'<b>{compliance_pct:.1f}%</b><br>Compliance',
                    x=0.5, y=0.5,
                    font_size=24,
                    font_color='white',
                    showarrow=False
                )
            ]
        )

        return fig

    def create_doctors_chart(self):
        """Cria gráfico de performance dos médicos"""
        if self.df_correlacionado is None or not self.medico_col:
            return None

        # Estatísticas por médico
        stats_medicos = self.df_correlacionado.groupby(self.medico_col).agg({
            'fora_do_prazo': ['count', 'sum'],
            'tempo_comunicacao_horas': 'mean'
        }).round(2)

        stats_medicos.columns = ['total_comunicados', 'fora_prazo', 'tempo_medio_h']
        stats_medicos['percentual_fora_prazo'] = (
            (stats_medicos['fora_prazo'] / stats_medicos['total_comunicados']) * 100
        ).round(1)

        # Top 10 médicos
        top_medicos = stats_medicos.nlargest(10, 'total_comunicados')

        # Criar gráfico de barras horizontal
        fig = go.Figure()

        # Adicionar barras com cores baseadas na performance
        colors = ['#27AE60' if pct <= 20 else '#F39C12' if pct <= 50 else '#E74C3C'
                 for pct in top_medicos['percentual_fora_prazo']]

        fig.add_trace(go.Bar(
            y=[str(nome).split()[-1] if str(nome).split() else str(nome) for nome in top_medicos.index],
            x=top_medicos['total_comunicados'],
            orientation='h',
            marker=dict(color=colors, opacity=0.8),
            text=[f"{row['total_comunicados']} casos<br>{row['percentual_fora_prazo']}% fora do prazo"
                  for _, row in top_medicos.iterrows()],
            textposition='auto',
            hovertemplate='<b>%{y}</b><br>Total: %{x} casos<br>Fora do prazo: %{customdata}%<extra></extra>',
            customdata=top_medicos['percentual_fora_prazo']
        ))

        fig.update_layout(
            title=dict(
                text='Top 10 Médicos - Total de Comunicados',
                x=0.5,
                font=dict(size=18, color='white')
            ),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(
                title='Total de Comunicados',
                gridcolor='rgba(255,255,255,0.1)',
                color='white'
            ),
            yaxis=dict(
                title='Médicos',
                gridcolor='rgba(255,255,255,0.1)',
                color='white'
            ),
            margin=dict(t=60, b=40, l=120, r=40),
            height=500
        )

        return fig

    def create_findings_chart(self):
        """Cria gráfico dos principais achados críticos"""
        if self.df_correlacionado is None or not self.achado_col:
            return None

        # Estatísticas por achado
        stats_achados = self.df_correlacionado.groupby(self.achado_col).agg({
            'fora_do_prazo': ['count', 'sum'],
            'tempo_comunicacao_horas': 'mean'
        }).round(2)

        stats_achados.columns = ['total_ocorrencias', 'fora_prazo', 'tempo_medio_h']
        stats_achados['percentual_fora_prazo'] = (
            (stats_achados['fora_prazo'] / stats_achados['total_ocorrencias']) * 100
        ).round(1)

        # Top 8 achados
        top_achados = stats_achados.nlargest(8, 'total_ocorrencias')

        # Simplificar nomes
        nomes_simplificados = []
        for nome in top_achados.index:
            nome_texto = str(nome)
            if ':' in nome_texto:
                partes = nome_texto.split(':')
                especialidade = partes[0].replace('Especialista em ', '').replace('Especialidade ', '')
                achado = partes[1].strip()[:40]
                nome_simp = f"{especialidade}: {achado}"
            else:
                nome_simp = nome_texto[:50]
            nomes_simplificados.append(nome_simp)

        fig = go.Figure()

        fig.add_trace(go.Bar(
            x=top_achados['total_ocorrencias'],
            y=nomes_simplificados,
            orientation='h',
            marker=dict(
                color=top_achados['percentual_fora_prazo'],
                colorscale='RdYlGn_r',
                colorbar=dict(
                    title=dict(
                        text="% Fora do Prazo",
                        font=dict(color='white')
                    ),
                    tickfont=dict(color='white')
                ),
                opacity=0.8
            ),
            text=[f"{freq} casos" for freq in top_achados['total_ocorrencias']],
            textposition='auto',
            hovertemplate='<b>%{y}</b><br>Total: %{x} casos<br>Fora do prazo: %{customdata}%<extra></extra>',
            customdata=top_achados['percentual_fora_prazo']
        ))

        fig.update_layout(
            title=dict(
                text='Top Achados Críticos por Frequência',
                x=0.5,
                font=dict(size=18, color='white')
            ),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(
                title='Total de Ocorrências',
                gridcolor='rgba(255,255,255,0.1)',
                color='white'
            ),
            yaxis=dict(
                title='',
                gridcolor='rgba(255,255,255,0.1)',
                color='white'
            ),
            margin=dict(t=60, b=40, l=280, r=40),
            height=500
        )

        return fig

    def create_time_distribution_chart(self):
        """Cria gráfico de distribuição dos tempos"""
        if self.df_correlacionado is None:
            return None

        # Criar bins para diferentes faixas de tempo
        bins = [float('-inf'), 0, 1, 6, 24, 168, float('inf')]
        labels = ['< 0h', '≤ 1h', '1-6h', '6-24h', '1-7 dias', '> 7 dias']

        self.df_correlacionado['faixa_tempo'] = pd.cut(
            self.df_correlacionado['tempo_comunicacao_horas'],
            bins=bins, labels=labels, right=False
        )

        faixa_counts = self.df_correlacionado['faixa_tempo'].value_counts(sort=False)

        colors = ['#3498DB', '#27AE60', '#F1C40F', '#E67E22', '#E74C3C', '#8E44AD']

        fig = go.Figure(data=[
            go.Bar(
                x=faixa_counts.index,
                y=faixa_counts.values,
                marker=dict(color=colors[:len(faixa_counts)], opacity=0.8),
                text=[f"{val}<br>({val/len(self.df_correlacionado)*100:.1f}%)"
                      for val in faixa_counts.values],
                textposition='auto',
                hovertemplate='<b>%{x}</b><br>Casos: %{y}<br>Percentual: %{customdata}%<extra></extra>',
                customdata=[f"{val/len(self.df_correlacionado)*100:.1f}" for val in faixa_counts.values]
            )
        ])

        fig.update_layout(
            title=dict(
                text='Distribuição por Faixa de Tempo',
                x=0.5,
                font=dict(size=18, color='white')
            ),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(
                title='Faixa de Tempo',
                gridcolor='rgba(255,255,255,0.1)',
                color='white'
            ),
            yaxis=dict(
                title='Número de Casos',
                gridcolor='rgba(255,255,255,0.1)',
                color='white'
            ),
            margin=dict(t=60, b=40, l=40, r=40),
            height=400
        )

        return fig

    def create_export_report(self):
        """Cria relatório para export"""
        if self.df_correlacionado is None:
            return None

        # Criar buffer para Excel
        output = io.BytesIO()
        export_df = self._format_date_columns(self.df_correlacionado)

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Sheet principal com dados correlacionados
            export_df.to_excel(writer, sheet_name='Dados_Correlacionados', index=False)

            # Sheet com estatísticas dos médicos
            if self.medico_col:
                stats_medicos = self.df_correlacionado.groupby(self.medico_col).agg({
                    'fora_do_prazo': ['count', 'sum'],
                    'tempo_comunicacao_horas': ['mean', 'median', 'max']
                }).round(2)
                stats_medicos.columns = ['total_comunicados', 'fora_prazo', 'tempo_medio_h', 'tempo_mediano_h', 'tempo_max_h']
                stats_medicos['percentual_fora_prazo'] = (
                    (stats_medicos['fora_prazo'] / stats_medicos['total_comunicados']) * 100
                ).round(2)
                stats_medicos.to_excel(writer, sheet_name='Estatisticas_Medicos')

            # Sheet com estatísticas dos achados
            if self.achado_col:
                stats_achados = self.df_correlacionado.groupby(self.achado_col).agg({
                    'fora_do_prazo': ['count', 'sum'],
                    'tempo_comunicacao_horas': ['mean', 'median', 'max']
                }).round(2)
                stats_achados.columns = ['total_ocorrencias', 'fora_prazo', 'tempo_medio_h', 'tempo_mediano_h', 'tempo_max_h']
                stats_achados['percentual_fora_prazo'] = (
                    (stats_achados['fora_prazo'] / stats_achados['total_ocorrencias']) * 100
                ).round(2)
                stats_achados.to_excel(writer, sheet_name='Estatisticas_Achados')

        output.seek(0)
        return output

    def _pdf_period_label(self):
        """Retorna o período atual do filtro para o relatório."""
        nomes_meses = {
            1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
            5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
            9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
        }
        filtro = st.session_state.get('current_filter', {}) if hasattr(st, 'session_state') else {}
        ano = filtro.get("ano")
        mes = filtro.get("mes")
        if ano and mes:
            return f"{nomes_meses.get(mes, mes)} {ano}"
        if ano:
            return f"Todo o ano {ano}"
        return "Periodo filtrado"

    def _pdf_detail_dataframe(self):
        """Monta a tabela detalhada de pacientes para o PDF."""
        columns = []
        rename_map = {}
        candidate_columns = [
            (self.same_col_achados, "SAME"),
            (self.nome_col_achados, "Paciente"),
            (self.data_col_achados, "Data Exame"),
            (self.desc_col_achados, "Procedimento"),
            (self.medico_col, "Medico Laudo"),
            (self.informado_por_col, "Informado Por"),
            (self.contato_col, "Contato"),
            (self.achado_col, "Achado Critico"),
            (self.data_sinalizacao_col, "Data Comunicacao"),
            (self.status_col, "Status A Laudar"),
            ("tempo_comunicacao_horas", "Tempo h"),
            ("fora_do_prazo", "Fora do prazo"),
        ]

        for source_col, display_col in candidate_columns:
            if source_col and source_col in self.df_correlacionado.columns and source_col not in columns:
                columns.append(source_col)
                rename_map[source_col] = display_col

        detail_df = self.df_correlacionado[columns].copy()
        detail_df = self._format_date_columns(detail_df)
        detail_df = detail_df.rename(columns=rename_map)

        if "Tempo h" in detail_df.columns:
            detail_df["Tempo h"] = pd.to_numeric(detail_df["Tempo h"], errors="coerce").map(
                lambda value: "" if pd.isna(value) else f"{value:.2f}"
            )
        if "Fora do prazo" in detail_df.columns:
            detail_df["Fora do prazo"] = detail_df["Fora do prazo"].map(lambda value: "SIM" if bool(value) else "Não")

        return detail_df

    def _wrap_pdf_value(self, value, width=24):
        """Quebra textos longos para caber na tabela do PDF."""
        text = "" if pd.isna(value) else str(value)
        if len(text) <= width:
            return text
        return "\n".join(textwrap.wrap(text, width=width, max_lines=3, placeholder="..."))

    def create_pdf_report(self):
        """Gera relatório PDF dark mode com métricas, gráficos e pacientes comunicados."""
        if self.df_correlacionado is None or len(self.df_correlacionado) == 0:
            return None

        from matplotlib.backends.backend_pdf import PdfPages
        import matplotlib.pyplot as plt

        bg = "#0e1117"
        panel = "#161b22"
        grid = "#30363d"
        text = "#f0f6fc"
        muted = "#8b949e"
        green = "#27AE60"
        red = "#E74C3C"
        yellow = "#F1C40F"
        orange = "#E67E22"
        blue = "#3498DB"
        purple = "#8E44AD"

        pdf_buffer = io.BytesIO()
        detail_df = self._pdf_detail_dataframe()
        total = len(self.df_correlacionado)
        no_prazo = int((~self.df_correlacionado['fora_do_prazo']).sum())
        fora_prazo = int(self.df_correlacionado['fora_do_prazo'].sum())
        compliance = (no_prazo / total * 100) if total else 0
        tempo_mediano = self.df_correlacionado['tempo_comunicacao_horas'].median()

        with PdfPages(pdf_buffer) as pdf:
            fig = plt.figure(figsize=(16, 9), facecolor=bg)
            gs = fig.add_gridspec(3, 4, height_ratios=[0.55, 1.4, 1.4], hspace=0.45, wspace=0.35)

            title_ax = fig.add_subplot(gs[0, :])
            title_ax.set_facecolor(bg)
            title_ax.axis("off")
            title_ax.text(
                0.01, 0.72,
                "CDI - Relatório de Achados Críticos",
                color=text,
                fontsize=26,
                fontweight="bold",
                ha="left",
            )
            title_ax.text(
                0.01, 0.25,
                f"{self._pdf_period_label()} | gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                color=muted,
                fontsize=12,
                ha="left",
            )

            metric_data = [
                ("Total", total, blue),
                ("No prazo", no_prazo, green),
                ("Fora do prazo", fora_prazo, red),
                ("Compliance", f"{compliance:.1f}%", green if compliance >= 80 else yellow if compliance >= 60 else red),
            ]
            for idx, (label, value, color) in enumerate(metric_data):
                ax = fig.add_subplot(gs[1, idx])
                ax.set_facecolor(panel)
                ax.set_xticks([])
                ax.set_yticks([])
                for spine in ax.spines.values():
                    spine.set_color(grid)
                ax.text(0.08, 0.72, str(value), color=color, fontsize=30, fontweight="bold", transform=ax.transAxes)
                ax.text(0.08, 0.35, label, color=text, fontsize=13, transform=ax.transAxes)
                if label == "Compliance":
                    ax.text(0.08, 0.15, f"Tempo mediano: {tempo_mediano:.2f}h", color=muted, fontsize=10, transform=ax.transAxes)

            ax_pie = fig.add_subplot(gs[2, 0:2])
            ax_pie.set_facecolor(panel)
            ax_pie.pie(
                [no_prazo, fora_prazo],
                labels=["No prazo", "Fora do prazo"],
                colors=[green, red],
                autopct="%1.1f%%",
                textprops={"color": text, "fontsize": 11},
                wedgeprops={"linewidth": 2, "edgecolor": bg},
            )
            ax_pie.set_title("Compliance de Comunicação", color=text, fontsize=15, pad=12)

            ax_time = fig.add_subplot(gs[2, 2:4])
            ax_time.set_facecolor(panel)
            bins = [float('-inf'), 0, 1, 6, 24, 168, float('inf')]
            labels = ['< 0h', '≤ 1h', '1-6h', '6-24h', '1-7 dias', '> 7 dias']
            time_bins = pd.cut(
                self.df_correlacionado['tempo_comunicacao_horas'],
                bins=bins,
                labels=labels,
                right=False,
            ).value_counts(sort=False)
            ax_time.bar(time_bins.index.astype(str), time_bins.values, color=[blue, green, yellow, orange, red, purple])
            ax_time.set_title("Distribuição por Tempo", color=text, fontsize=15, pad=12)
            ax_time.tick_params(colors=text, labelrotation=20)
            ax_time.grid(axis="y", color=grid, alpha=0.45)
            for spine in ax_time.spines.values():
                spine.set_color(grid)

            pdf.savefig(fig, facecolor=bg, bbox_inches="tight")
            plt.close(fig)

            fig = plt.figure(figsize=(16, 9), facecolor=bg)
            gs = fig.add_gridspec(1, 2, wspace=0.35)

            ax_doc = fig.add_subplot(gs[0, 0])
            ax_doc.set_facecolor(panel)
            if self.medico_col:
                stats_medicos = self.df_correlacionado.groupby(self.medico_col).agg({
                    'fora_do_prazo': ['count', 'sum']
                })
                stats_medicos.columns = ['total', 'fora']
                stats_medicos['percentual_fora'] = stats_medicos['fora'] / stats_medicos['total'] * 100
                top_medicos = stats_medicos.nlargest(10, 'total').sort_values('total')
                colors_doc = [red if pct > 50 else yellow if pct > 20 else green for pct in top_medicos['percentual_fora']]
                ax_doc.barh([str(v).split()[-1] for v in top_medicos.index], top_medicos['total'], color=colors_doc)
            ax_doc.set_title("Top Médicos por Comunicados", color=text, fontsize=15, pad=12)
            ax_doc.tick_params(colors=text)
            ax_doc.grid(axis="x", color=grid, alpha=0.45)
            for spine in ax_doc.spines.values():
                spine.set_color(grid)

            ax_find = fig.add_subplot(gs[0, 1])
            ax_find.set_facecolor(panel)
            if self.achado_col:
                stats_achados = self.df_correlacionado.groupby(self.achado_col).agg({
                    'fora_do_prazo': ['count', 'sum']
                })
                stats_achados.columns = ['total', 'fora']
                top_achados = stats_achados.nlargest(8, 'total').sort_values('total')
                labels_find = [
                    self._wrap_pdf_value(str(label).split(':')[-1].strip(), width=30)
                    for label in top_achados.index
                ]
                ax_find.barh(labels_find, top_achados['total'], color=purple)
            ax_find.set_title("Achados Críticos Mais Frequentes", color=text, fontsize=15, pad=12)
            ax_find.tick_params(colors=text)
            ax_find.grid(axis="x", color=grid, alpha=0.45)
            for spine in ax_find.spines.values():
                spine.set_color(grid)

            pdf.savefig(fig, facecolor=bg, bbox_inches="tight")
            plt.close(fig)

            rows_per_page = 9
            visible_cols = [col for col in [
                "SAME", "Paciente", "Data Exame", "Procedimento", "Medico Laudo",
                "Data Comunicacao", "Tempo h", "Fora do prazo"
            ] if col in detail_df.columns]

            for start in range(0, len(detail_df), rows_per_page):
                page_df = detail_df.iloc[start:start + rows_per_page][visible_cols].copy()
                for col in page_df.columns:
                    wrap_width = 34 if col in {"Paciente", "Procedimento", "Medico Laudo"} else 18
                    page_df[col] = page_df[col].map(lambda value, width=wrap_width: self._wrap_pdf_value(value, width))

                fig, ax = plt.subplots(figsize=(16, 9), facecolor=bg)
                ax.set_facecolor(bg)
                ax.axis("off")
                ax.text(
                    0.0, 1.04,
                    f"Pacientes Comunicados - {self._pdf_period_label()}",
                    color=text,
                    fontsize=18,
                    fontweight="bold",
                    transform=ax.transAxes,
                )
                ax.text(
                    0.0, 1.0,
                    f"Página {start // rows_per_page + 1} | registros {start + 1}-{min(start + rows_per_page, len(detail_df))} de {len(detail_df)}",
                    color=muted,
                    fontsize=10,
                    transform=ax.transAxes,
                )

                table = ax.table(
                    cellText=page_df.values,
                    colLabels=page_df.columns,
                    loc="center",
                    cellLoc="left",
                    colLoc="left",
                )
                table.auto_set_font_size(False)
                table.set_fontsize(7.4)
                table.scale(1, 2.25)

                prazo_col_idx = page_df.columns.get_loc("Fora do prazo") if "Fora do prazo" in page_df.columns else None
                for (row_idx, col_idx), cell in table.get_celld().items():
                    cell.set_edgecolor(grid)
                    if row_idx == 0:
                        cell.set_facecolor("#1f6feb")
                        cell.get_text().set_color("white")
                        cell.get_text().set_weight("bold")
                        continue

                    base_color = panel if row_idx % 2 else "#111820"
                    if prazo_col_idx is not None and col_idx == prazo_col_idx:
                        value = page_df.iloc[row_idx - 1, col_idx]
                        if str(value).strip().upper() == "SIM":
                            base_color = red
                    cell.set_facecolor(base_color)
                    cell.get_text().set_color(text)

                pdf.savefig(fig, facecolor=bg, bbox_inches="tight")
                plt.close(fig)

        pdf_buffer.seek(0)
        return pdf_buffer

    def render_analysis_results(self):
        """Renderiza os resultados da análise"""
        if self.df_correlacionado is None or len(self.df_correlacionado) == 0:
            st.warning("📊 Nenhum dado para análise. Faça o upload das planilhas primeiro.")
            return

        # Métricas principais
        st.markdown("## 📊 Visão Geral")
        self.render_metrics_overview()

        # Gráficos principais
        col1, col2 = st.columns(2)

        with col1:
            fig_compliance = self.create_compliance_chart()
            if fig_compliance:
                st.plotly_chart(fig_compliance, use_container_width=True)

        with col2:
            fig_time = self.create_time_distribution_chart()
            if fig_time:
                st.plotly_chart(fig_time, use_container_width=True)

        # Gráficos detalhados
        st.markdown("## 👨‍⚕️ Análise por Médicos")
        fig_doctors = self.create_doctors_chart()
        if fig_doctors:
            st.plotly_chart(fig_doctors, use_container_width=True)

        st.markdown("## 🔍 Análise por Achados Críticos")
        fig_findings = self.create_findings_chart()
        if fig_findings:
            st.plotly_chart(fig_findings, use_container_width=True)

        # Tabela de dados
        with st.expander("📋 Dados Detalhados", expanded=False):
            # Criar lista de colunas disponíveis
            cols_to_show = []
            if self.same_col_achados:
                cols_to_show.append(self.same_col_achados)
            if self.nome_col_achados:
                cols_to_show.append(self.nome_col_achados)
            if self.medico_col:
                cols_to_show.append(self.medico_col)
            if self.contato_col:
                cols_to_show.append(self.contato_col)
            if self.informado_por_col:
                cols_to_show.append(self.informado_por_col)
            if self.achado_col:
                cols_to_show.append(self.achado_col)
            if self.data_sinalizacao_col:
                cols_to_show.append(self.data_sinalizacao_col)
            if self.status_col:
                cols_to_show.append(self.status_col)
            cols_to_show.extend([
                'tempo_comunicacao_horas',
                'fora_do_prazo',
                'tempo_negativo',
                'match_confidence',
                'match_score',
                'match_procedure_similarity',
                'match_time_delta_min',
            ])

            # Filtrar apenas colunas que existem no DataFrame
            cols_to_show = [col for col in cols_to_show if col in self.df_correlacionado.columns]

            if cols_to_show:
                df_display = self._format_date_columns(self.df_correlacionado[cols_to_show])
                st.dataframe(
                    df_display.sort_values('tempo_comunicacao_horas', ascending=False),
                    use_container_width=True
                )

        review_df = getattr(self, 'df_revisao_correlacao', None)
        if review_df is not None and not review_df.empty:
            with st.expander("⚠️ Registros para Revisão de Correlação", expanded=False):
                review_cols = [
                    self.same_col_achados,
                    self.nome_col_achados,
                    self.data_col_achados,
                    self.desc_col_achados,
                    'match_status',
                    'match_candidates',
                    'match_score',
                    'match_procedure_similarity',
                    'match_candidate_procedure',
                    'match_candidate_datetime',
                ]
                review_cols = [col for col in review_cols if col and col in review_df.columns]
                review_display = self._format_date_columns(review_df[review_cols])
                st.dataframe(review_display, use_container_width=True)

                review_output = io.BytesIO()
                with pd.ExcelWriter(review_output, engine='openpyxl') as writer:
                    self._format_date_columns(review_df).to_excel(
                        writer,
                        sheet_name='Revisao_Correlacao',
                        index=False,
                    )
                review_output.seek(0)
                st.download_button(
                    label="⬇️ Baixar Revisão de Correlação",
                    data=review_output,
                    file_name=f"revisao_correlacao_achados_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # Export
        st.markdown("## 📥 Export de Relatórios")
        col1, col2, col3 = st.columns([1, 1, 2])

        with col1:
            if st.button("📊 Gerar Relatório Excel", type="primary"):
                excel_data = self.create_export_report()
                if excel_data:
                    st.download_button(
                        label="⬇️ Download Excel",
                        data=excel_data,
                        file_name=f"relatorio_achados_criticos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        with col2:
            if st.button("📄 Gerar PDF Dark Mode", type="primary"):
                pdf_data = self.create_pdf_report()
                if pdf_data:
                    st.download_button(
                        label="⬇️ Download PDF",
                        data=pdf_data,
                        file_name=f"relatorio_achados_criticos_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                        mime="application/pdf"
                    )

def main():
    """Função principal do dashboard"""
    # Inicializar session_state
    if 'df_correlacionado' not in st.session_state:
        st.session_state.df_correlacionado = None
    if 'df_correlacionado_full' not in st.session_state:
        st.session_state.df_correlacionado_full = None
    if 'processed' not in st.session_state:
        st.session_state.processed = False
    if 'ris_extracted_df' not in st.session_state:
        st.session_state.ris_extracted_df = pd.DataFrame(columns=["Campo", "Valor"])
    if 'ris_image_signature' not in st.session_state:
        st.session_state.ris_image_signature = None
    if 'ris_debug_payload' not in st.session_state:
        st.session_state.ris_debug_payload = None

    dashboard = DashboardAchadosCriticos()

    # Restaurar dados do session_state
    if st.session_state.df_correlacionado_full is not None:
        dashboard.df_correlacionado = st.session_state.df_correlacionado_full.copy()
    elif st.session_state.df_correlacionado is not None:
        dashboard.df_correlacionado = st.session_state.df_correlacionado.copy()

    # Header
    dashboard.render_header()

    # Sidebar
    uploaded_achados, uploaded_status = dashboard.render_sidebar()

    # Controles na sidebar
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🚀 Processar Dados")

    process_button = st.sidebar.button("🔄 Processar Planilhas", type="primary")

    # Se ambos os arquivos foram carregados
    if uploaded_achados is not None and uploaded_status is not None:
        if process_button:
            with st.spinner("🔄 Processando dados..."):
                # Carregar dados
                load_success = dashboard.load_data(uploaded_achados, uploaded_status)
                if load_success:
                    st.sidebar.success("✅ Planilhas carregadas")
                    # Correlacionar
                    if dashboard.correlate_data():
                        st.sidebar.success("✅ Dados correlacionados")

                        # Calcular tempos
                        if dashboard.calculate_times():
                            st.sidebar.success("✅ Tempos calculados")

                            # Salvar no session_state
                            st.session_state.df_correlacionado_full = dashboard.df_correlacionado.copy()
                            st.session_state.df_correlacionado = dashboard.df_correlacionado.copy()
                            st.session_state.processed = True
                            st.session_state.match_summary = getattr(dashboard, 'match_summary', {})
                            st.session_state.time_summary = getattr(dashboard, 'time_summary', {})
                            st.session_state.df_revisao_correlacao = getattr(
                                dashboard,
                                'df_revisao_correlacao',
                                pd.DataFrame(),
                            )

                            # Copiar atributos necessários
                            st.session_state.medico_col = dashboard.medico_col
                            st.session_state.achado_col = dashboard.achado_col
                            st.session_state.same_col_achados = dashboard.same_col_achados
                            st.session_state.nome_col_achados = dashboard.nome_col_achados
                            st.session_state.nome_col_status = dashboard.nome_col_status
                            st.session_state.data_col_achados = dashboard.data_col_achados
                            st.session_state.data_col_status = dashboard.data_col_status
                            st.session_state.data_sinalizacao_col = dashboard.data_sinalizacao_col
                            st.session_state.status_col = dashboard.status_col
                            st.session_state.desc_col_achados = dashboard.desc_col_achados
                            st.session_state.desc_col_status = dashboard.desc_col_status
                            st.session_state.modalidade_col_achados = dashboard.modalidade_col_achados
                            st.session_state.modalidade_col_status = dashboard.modalidade_col_status
                            st.session_state.contato_col = dashboard.contato_col
                            st.session_state.informado_por_col = dashboard.informado_por_col

                            st.rerun()
                        else:
                            st.sidebar.error("❌ Erro no cálculo de tempos")
                    else:
                        st.sidebar.error("❌ Erro na correlação")
                else:
                    st.sidebar.error("❌ Erro ao carregar dados")

    dashboard_tab, ris_tab = st.tabs(["Dashboard Analitico", "RIS OCR / Email"])

    with dashboard_tab:
        # Restaurar atributos do session_state para renderização
        if st.session_state.processed:
            full_df = st.session_state.get('df_correlacionado_full')
            dashboard.df_correlacionado = (
                full_df.copy() if full_df is not None else st.session_state.df_correlacionado.copy()
            )
            review_df = st.session_state.get('df_revisao_correlacao')
            dashboard.df_revisao_correlacao = review_df.copy() if review_df is not None else None
            dashboard.medico_col = st.session_state.get('medico_col')
            dashboard.achado_col = st.session_state.get('achado_col')
            dashboard.same_col_achados = st.session_state.get('same_col_achados')
            dashboard.nome_col_achados = st.session_state.get('nome_col_achados')
            dashboard.nome_col_status = st.session_state.get('nome_col_status')
            dashboard.data_col_achados = st.session_state.get('data_col_achados')
            dashboard.data_col_status = st.session_state.get('data_col_status')
            dashboard.data_sinalizacao_col = st.session_state.get('data_sinalizacao_col')
            dashboard.status_col = st.session_state.get('status_col')
            dashboard.desc_col_achados = st.session_state.get('desc_col_achados')
            dashboard.desc_col_status = st.session_state.get('desc_col_status')
            dashboard.modalidade_col_achados = st.session_state.get('modalidade_col_achados')
            dashboard.modalidade_col_status = st.session_state.get('modalidade_col_status')
            dashboard.contato_col = st.session_state.get('contato_col')
            dashboard.informado_por_col = st.session_state.get('informado_por_col')

            match_summary = st.session_state.get('match_summary')
            time_summary = st.session_state.get('time_summary')
            if match_summary:
                total_revisar = match_summary.get('sem_same', 0) + match_summary.get('sem_confianca', 0)
                st.sidebar.markdown("---")
                st.sidebar.markdown("### 🔎 Qualidade da Correlação")
                st.sidebar.caption(
                    f"Alta: {match_summary.get('alta', 0)} | "
                    f"Média: {match_summary.get('media', 0)} | "
                    f"Revisar: {total_revisar}"
                )
            if time_summary:
                st.sidebar.caption(
                    f"Calculados: {time_summary.get('calculados', 0)} | "
                    f"Sem status: {time_summary.get('sem_status', 0)} | "
                    f"Datas inválidas: {time_summary.get('datas_invalidas', 0)}"
                )

            # Renderizar filtro de data
            ano_filtro, mes_filtro = dashboard.render_date_filter()

            # Aplicar filtro de data se selecionado
            if ano_filtro is not None and dashboard.data_sinalizacao_col:
                registros_filtrados, registros_totais = dashboard.apply_date_filter(ano_filtro, mes_filtro)
                st.session_state.df_correlacionado = dashboard.df_correlacionado.copy()
                st.session_state.current_filter = {
                    "ano": ano_filtro,
                    "mes": mes_filtro,
                    "registros_filtrados": registros_filtrados,
                    "registros_totais": registros_totais,
                }
                st.sidebar.success(
                    f"✅ {registros_filtrados} de {registros_totais} registro(s) no período selecionado"
                )

        # Renderizar resultados
        dashboard.render_analysis_results()

    with ris_tab:
        dashboard.render_ris_ocr_email_tab()

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; opacity: 0.7; padding: 20px;'>
        🏥 Sistema de Análise de Achados Críticos - CDI<br>
        Desenvolvido com ❤️ usando Streamlit + Plotly
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
