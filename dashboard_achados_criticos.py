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

# Configuracao fixa do servidor SMTP. O admin deve preencher esses valores.
ADMIN_SMTP_CONFIG = {
    "host": "smtp.seu-servidor.local",
    "port": 587,
    "use_tls": True,
    "username": "no-reply@hospital.local",
    "password": "ALTERAR_PELO_ADMIN",
    "sender_email": "no-reply@hospital.local",
    "sender_name": "CDI - Achados Criticos",
    "tesseract_cmd": os.environ.get("TESSERACT_CMD", "/usr/bin/tesseract"),
}

# Regioes relativas ao dialogo "Resultado Critico".
RIS_DIALOG_FIELD_REGIONS = [
    {"field": "Resultado Crítico", "box": (0.17, 0.12, 0.40, 0.20), "multiline": False},
    {"field": "Contato", "box": (0.16, 0.22, 0.56, 0.31), "multiline": False},
    {"field": "Contato com (Sucesso)", "box": (0.78, 0.22, 0.97, 0.31), "multiline": False},
    {"field": "Data e Hora", "box": (0.18, 0.42, 0.55, 0.54), "multiline": False},
    {"field": "Observações", "box": (0.16, 0.56, 0.95, 0.89), "multiline": True},
]

RIS_DIAGNOSIS_RELATIVE_BOX = (-0.07, -0.19, 0.43, -0.07)

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
        "psm": 7,
        "scale": 6,
        "whitelist": "SsIiMmNnAaOoÃãÕõ",
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

    def _format_date_columns(self, df):
        """Formata colunas de data para DD/MM/AAAA apenas para exibição/export."""
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
                formatted = parsed.dt.strftime('%d/%m/%Y')
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

        if field_name == "Contato com (Sucesso)":
            lowered = cleaned.lower()
            if "sim" in lowered:
                return "Sim"
            if "nao" in lowered or "não" in lowered:
                return "Não"

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
        """Valida a configuracao fixa de SMTP."""
        required_keys = ["host", "port", "sender_email"]
        missing = [key for key in required_keys if not ADMIN_SMTP_CONFIG.get(key)]
        if missing:
            return False, f"Configuracao SMTP incompleta: {', '.join(missing)}"

        if ADMIN_SMTP_CONFIG.get("password") == "ALTERAR_PELO_ADMIN":
            return False, "Senha SMTP ainda nao foi configurada pelo admin."

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
        # Função auxiliar para normalizar strings
        def normalize(text):
            if not isinstance(text, str):
                return text
            return text.lower().replace('ç', 'c').replace('ã', 'a').replace('á', 'a').replace('í', 'i').replace('ó', 'o')

        # Encontrar coluna SAME
        self.same_col_achados = None
        self.same_col_status = None
        for col in self.df_achados.columns:
            if isinstance(col, str) and 'same' in normalize(col):
                self.same_col_achados = col
                break

        for col in self.df_status.columns:
            if isinstance(col, str) and 'same' in normalize(col):
                self.same_col_status = col
                break

        # Encontrar coluna Nome do Paciente
        self.nome_col_achados = None
        self.nome_col_status = None
        for col in self.df_achados.columns:
            col_norm = normalize(col)
            if isinstance(col, str) and 'nome' in col_norm and 'paciente' in col_norm:
                self.nome_col_achados = col
                break

        for col in self.df_status.columns:
            col_norm = normalize(col)
            if isinstance(col, str) and 'nome' in col_norm and 'paciente' in col_norm:
                self.nome_col_status = col
                break

        # Encontrar coluna Data Exame
        self.data_col_achados = None
        self.data_col_status = None
        for col in self.df_achados.columns:
            col_norm = normalize(col)
            if isinstance(col, str) and 'data' in col_norm and 'exame' in col_norm:
                self.data_col_achados = col
                break

        for col in self.df_status.columns:
            col_norm = normalize(col)
            if isinstance(col, str) and 'data' in col_norm and ('prescri' in col_norm or 'hora' in col_norm):
                self.data_col_status = col
                break

        # Encontrar coluna Descrição Procedimento
        self.desc_col_achados = None
        self.desc_col_status = None
        for col in self.df_achados.columns:
            col_norm = normalize(col)
            if isinstance(col, str) and 'descri' in col_norm and 'proced' in col_norm:
                self.desc_col_achados = col
                break

        for col in self.df_status.columns:
            col_norm = normalize(col)
            if isinstance(col, str) and 'descri' in col_norm and 'proced' in col_norm:
                self.desc_col_status = col
                break

        # Encontrar STATUS_ALAUDAR
        self.status_col = None
        for col in self.df_status.columns:
            if isinstance(col, str) and 'status' in normalize(col) and 'laudar' in normalize(col):
                self.status_col = col
                break

        # Encontrar Data Sinalização
        self.data_sinalizacao_col = None
        for col in self.df_achados.columns:
            col_norm = normalize(col)
            if isinstance(col, str) and 'data' in col_norm and 'sinalizacao' in col_norm:
                self.data_sinalizacao_col = col
                break

        # Encontrar coluna Medico Laudo
        self.medico_col = None
        for col in self.df_achados.columns:
            col_norm = normalize(col)
            if isinstance(col, str) and 'medico' in col_norm and 'laudo' in col_norm:
                self.medico_col = col
                break

        # Encontrar coluna Achado Crítico
        self.achado_col = None
        for col in self.df_achados.columns:
            col_norm = normalize(col)
            if isinstance(col, str) and 'achado' in col_norm and 'critico' in col_norm:
                self.achado_col = col
                break

        # Encontrar coluna Contato (médico que recebeu o contato)
        self.contato_col = None
        for col in self.df_achados.columns:
            col_norm = normalize(col)
            if isinstance(col, str) and 'contato' in col_norm and 'sucesso' not in col_norm:
                self.contato_col = col
                break

        # Encontrar coluna Informado Por (médico que fez a comunicação)
        self.informado_por_col = None
        for col in self.df_achados.columns:
            col_norm = normalize(col)
            if isinstance(col, str) and 'informado' in col_norm and 'por' in col_norm:
                self.informado_por_col = col
                break

        return (self.same_col_achados is not None and
                self.same_col_status is not None)

    def correlate_data(self):
        """Correlaciona as planilhas usando múltiplos critérios"""
        if self.df_achados is None or self.df_status is None:
            return False

        try:
            # Identificar colunas
            if not self.identify_columns():
                st.error("❌ Não foi possível identificar as colunas necessárias")
                return False

            correlacoes = []

            for idx, achado in self.df_achados.iterrows():
                same_achado = achado[self.same_col_achados]
                nome_achado = achado.get(self.nome_col_achados, '') if self.nome_col_achados else ''
                data_exame_achado = achado.get(self.data_col_achados, None) if self.data_col_achados else None
                descricao_achado = achado.get(self.desc_col_achados, '') if self.desc_col_achados else ''

                # Buscar exames correspondentes
                exames_same = self.df_status[self.df_status[self.same_col_status] == same_achado]

                if len(exames_same) == 0:
                    correlacao = achado.copy()
                    correlacoes.append(correlacao)
                    continue

                # Filtrar por nome
                if pd.notna(nome_achado) and isinstance(nome_achado, str) and self.nome_col_status:
                    exames_nome = exames_same[
                        exames_same[self.nome_col_status].str.contains(
                            nome_achado.split()[0], case=False, na=False
                        ) |
                        exames_same[self.nome_col_status].str.contains(
                            nome_achado.split()[-1], case=False, na=False
                        )
                    ]
                else:
                    exames_nome = exames_same

                # Filtrar por data
                if pd.notna(data_exame_achado) and self.data_col_status:
                    try:
                        data_achado_dt = self._parse_datetime_value(data_exame_achado)
                        if pd.isna(data_achado_dt):
                            exames_finais = exames_nome
                        else:
                            data_achado_str = data_achado_dt.strftime('%d/%m/%Y')

                            exames_data = []
                            for _, exame in exames_nome.iterrows():
                                if pd.notna(exame.get(self.data_col_status)):
                                    try:
                                        data_exame_dt = self._parse_datetime_value(exame[self.data_col_status])
                                        if pd.isna(data_exame_dt):
                                            continue
                                        data_exame_str = data_exame_dt.strftime('%d/%m/%Y')
                                        if data_achado_str == data_exame_str:
                                            exames_data.append(exame)
                                    except:
                                        continue

                            if len(exames_data) > 0:
                                exames_finais = pd.DataFrame(exames_data)
                            else:
                                exames_finais = exames_nome
                    except:
                        exames_finais = exames_nome
                else:
                    exames_finais = exames_nome

                # Selecionar melhor exame
                if len(exames_finais) > 0:
                    if self.status_col:
                        exames_com_status = exames_finais[exames_finais[self.status_col].notna()]
                        melhor_exame = exames_com_status.iloc[0] if len(exames_com_status) > 0 else exames_finais.iloc[0]
                    else:
                        melhor_exame = exames_finais.iloc[0]

                    correlacao = achado.copy()
                    for col in melhor_exame.index:
                        if col != self.same_col_status:
                            correlacao[col] = melhor_exame[col]

                    correlacoes.append(correlacao)

            self.df_correlacionado = pd.DataFrame(correlacoes)
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

            self.df_correlacionado = df_com_status[registros_validos].copy()

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
            (self.df_correlacionado[self.medico_col].str.strip() != '')
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
                    text=f'<b>{(no_prazo/(no_prazo+fora_prazo)*100):.1f}%</b><br>Compliance',
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
            y=[nome.split()[-1] for nome in top_medicos.index],  # Só sobrenome
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
            if ':' in nome:
                partes = nome.split(':')
                especialidade = partes[0].replace('Especialista em ', '').replace('Especialidade ', '')
                achado = partes[1].strip()[:40]
                nome_simp = f"{especialidade}: {achado}"
            else:
                nome_simp = nome[:50]
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
        bins = [0, 1, 6, 24, 168, float('inf')]
        labels = ['≤ 1h', '1-6h', '6-24h', '1-7 dias', '> 7 dias']

        self.df_correlacionado['faixa_tempo'] = pd.cut(
            self.df_correlacionado['tempo_comunicacao_horas'],
            bins=bins, labels=labels, right=False
        )

        faixa_counts = self.df_correlacionado['faixa_tempo'].value_counts()

        colors = ['#27AE60', '#F1C40F', '#E67E22', '#E74C3C', '#8E44AD']

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
            cols_to_show.extend(['tempo_comunicacao_horas', 'fora_do_prazo'])

            # Filtrar apenas colunas que existem no DataFrame
            cols_to_show = [col for col in cols_to_show if col in self.df_correlacionado.columns]

            if cols_to_show:
                df_display = self._format_date_columns(self.df_correlacionado[cols_to_show])
                st.dataframe(
                    df_display.sort_values('tempo_comunicacao_horas', ascending=False),
                    use_container_width=True
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

def main():
    """Função principal do dashboard"""
    # Inicializar session_state
    if 'df_correlacionado' not in st.session_state:
        st.session_state.df_correlacionado = None
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
    if st.session_state.df_correlacionado is not None:
        dashboard.df_correlacionado = st.session_state.df_correlacionado

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
                            st.session_state.df_correlacionado = dashboard.df_correlacionado
                            st.session_state.processed = True

                            # Copiar atributos necessários
                            st.session_state.medico_col = dashboard.medico_col
                            st.session_state.achado_col = dashboard.achado_col
                            st.session_state.same_col_achados = dashboard.same_col_achados
                            st.session_state.nome_col_achados = dashboard.nome_col_achados
                            st.session_state.data_sinalizacao_col = dashboard.data_sinalizacao_col
                            st.session_state.status_col = dashboard.status_col
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
            dashboard.df_correlacionado = st.session_state.df_correlacionado
            dashboard.medico_col = st.session_state.get('medico_col')
            dashboard.achado_col = st.session_state.get('achado_col')
            dashboard.same_col_achados = st.session_state.get('same_col_achados')
            dashboard.nome_col_achados = st.session_state.get('nome_col_achados')
            dashboard.data_sinalizacao_col = st.session_state.get('data_sinalizacao_col')
            dashboard.status_col = st.session_state.get('status_col')
            dashboard.contato_col = st.session_state.get('contato_col')
            dashboard.informado_por_col = st.session_state.get('informado_por_col')

            # Renderizar filtro de data
            ano_filtro, mes_filtro = dashboard.render_date_filter()

            # Aplicar filtro de data se selecionado
            if ano_filtro is not None and dashboard.data_sinalizacao_col:
                datas = dashboard._parse_datetime_series(dashboard.df_correlacionado[dashboard.data_sinalizacao_col])

                # Filtrar por ano
                mask_ano = datas.dt.year == ano_filtro

                # Filtrar por mês se especificado
                if mes_filtro is not None:
                    mask_mes = datas.dt.month == mes_filtro
                    mask_final = mask_ano & mask_mes
                else:
                    mask_final = mask_ano

                # Aplicar filtro
                dashboard.df_correlacionado = dashboard.df_correlacionado[mask_final].copy()

                # Mostrar contador de registros
                st.sidebar.success(f"✅ {len(dashboard.df_correlacionado)} registros no período selecionado")

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
