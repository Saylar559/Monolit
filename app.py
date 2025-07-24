import streamlit as st  # Для создания веб-интерфейса
import pandas as pd     # Для работы с табличными данными
import io               # Для работы с потоками ввода-вывода
import os               # Для работы с операционной системой (файлы, пути)
from datetime import datetime  # Для работы с датой и временем
import base64           # Для кодирования/декодирования base64 (например, для файлов)

# Добавление пользовательских стилей CSS для оформления интерфейса
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
    body {
        font-family: 'Roboto', sans-serif;
        background-color: #f5f5f5;
    }
    .title {
        font-size: 2.5em;
        font-weight: 700;
        color: #272727;
        text-align: center;
        margin-bottom: 0.5em;
    }
    .subtitle {
        font-size: 1.5em;
        font-weight: 400;
        color: #4a4a4a;
        text-align: center;
        margin-bottom: 1em;
    }
.download-link {
    background-color: #4CAF50;
    color: white;
    padding: 10px 50px;
    border: none;
    border-radius: 10px;
    font-size: 16px;
    cursor: pointer;
    transition: background-color 0.3s ease;
    float: right;
}
.st-emotion-cache-1nwdr1w a {
    color: white;
    text-decoration: initial;
}

.download-link:hover {
 background-color: #45a049; /* Темнее зеленый при наведении */
}


.download-link:active {
    transform: translateY(0); /* Reset lift on click */
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); /* Subtle shadow on click */
}
    .success-message {
        color: #388e3c;
        font-weight: 500;
        margin-top: 10px;
        text-align: right;
    }
    #excel_drop_zone {
        border: 2px dashed #4a4a4a;
        padding: 20px;
        text-align: center;
        background-color: #ffffff;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# Функция для создания ссылки на скачивание файла
def get_binary_file_downloader_html(bin_file, file_label='File'):
    bin_str = base64.b64encode(bin_file).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{file_label}" class="download-link">СКАЧАТЬ</a>'
    return href

# Функция для парсинга дат из различных форматов
def parse_date(date_str):
    if pd.isna(date_str):
        return pd.NaT
    
    try:
        date_num = float(date_str)
        return pd.to_datetime('1899-12-30') + pd.to_timedelta(date_num, unit='D')
    except (ValueError, TypeError):
        pass
    
    date_formats = [
        '%d.%m.%Y', '%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y', 
        '%Y.%m.%d', '%d %b %Y', '%d %B %Y', '%Y%m%d'
    ]
    
    for fmt in date_formats:
        try:
            return pd.to_datetime(date_str, format=fmt)
        except (ValueError, TypeError):
            continue
    
    return pd.NaT

# Функция для форматирования суммы в требуемом формате
def format_amount(amount):
    return f"{amount:,.2f}".replace(",", " ").replace(".", ",") + " руб."

# Функция для анализа Excel-файлов
def analyze_excel_files(excel_files, year=None, month=None, filter_by_period=True):
    results = []
    errors = []
    target_year_month = f"{year}-{month:02d}" if filter_by_period else None
    permit_mapping = {
        '91-RU93308000-2132-2022': 'Поступления на счет Эскроу "Горизонт 1"',
        '91-RU93308000-2775-2023': 'Поступления на счет Эскроу "Горизонт 2"',
        '91-RU93308000-3161-2023': 'Поступления на счет Эскроу "Горизонт 3"'
    }

    for excel_file in excel_files:
        try:
            xl = pd.read_excel(excel_file, sheet_name=None, skiprows=6)
            file_has_data = False
            file_name_base = excel_file.name.split('.')[0]
            default_file_name = f"Поступления на счет Эскроу {file_name_base}"

            for sheet_name, df in xl.items():
                if 'лист' in sheet_name.lower():
                    continue
                if df.empty:
                    continue

                df.columns = df.columns.str.strip()
                if 'Сумма операции' in df.columns:
                    df = df.rename(columns={'Сумма операции': 'Сумма поступления / списания, руб'})
                if 'Дата операции' in df.columns:
                    df = df.rename(columns={'Дата операции': 'Дата поступления / списания'})

                if 'Сумма поступления / списания, руб' not in df.columns:
                    errors.append({
                        'Название обьекта': f"{default_file_name} (лист {sheet_name})",
                        'Причина': 'Отсутствует столбец "Сумма поступления / списания, руб"'
                    })
                    continue

                df['Сумма'] = pd.to_numeric(df['Сумма поступления / списания, руб'], errors='coerce')
                df = df.dropna(subset=['Сумма'])

                if 'Дата поступления / списания' in df.columns:
                    df['Дата'] = df['Дата поступления / списания'].apply(parse_date)
                    if df['Дата'].isna().all():
                        errors.append({
                            'Название обьекта': f"{default_file_name} (лист {sheet_name})",
                            'Причина': 'Не удалось распознать даты в столбце "Дата поступления / списания"'
                        })
                        continue
                else:
                    errors.append({
                        'Название обьекта': f"{default_file_name} (лист {sheet_name})",
                        'Причина': 'Отсутствует столбец "Дата поступления / списания"'
                    })
                    continue

                df = df.dropna(subset=['Дата'])
                if df.empty:
                    continue

                df['Дата_отображения'] = df['Дата'].dt.strftime('%d.%m.%Y')
                if filter_by_period and target_year_month:
                    df = df[df['Дата'].dt.strftime("%Y-%m") == target_year_month]

                df_positive = df[df['Сумма'] > 0]
                file_has_data = True

                if 'Разрешение на строительство' in df.columns:
                    df_positive['Название обьекта'] = df_positive['Разрешение на строительство'].map(permit_mapping).fillna(default_file_name)
                    grouped = df_positive.groupby('Название обьекта').agg({
                        'Сумма': 'sum',
                        'Название обьекта': 'count'
                    }).rename(columns={'Название обьекта': 'Количество операций'})
                    grouped = grouped.reset_index()
                    for _, row in grouped.iterrows():
                        results.append({
                            'Название обьекта': row['Название обьекта'],
                            'Сумма': format_amount(row['Сумма']),
                        })
                else:
                    total_sum = df_positive['Сумма'].sum()
                    results.append({
                        'Название обьекта': default_file_name,
                        'Сумма': format_amount(total_sum),
                    })

            if not file_has_data:
                results.append({
                    'Название обьекта': default_file_name,
                    'Сумма': format_amount(0),
                })

        except Exception as e:
            errors.append({
                'Название обьекта': default_file_name,
                'Причина': f'Ошибка при обработке: {str(e)}'
            })

    if results:
        result_df = pd.DataFrame(results)
    else:
        result_df = pd.DataFrame(columns=['Название обьекта', 'Сумма'])

    error_df = pd.DataFrame(errors) if errors else pd.DataFrame(columns=['Название обьекта', 'Причина'])
    return result_df, error_df

# Основной интерфейс Streamlit
st.markdown('<h1 class="title">Анализ финансовых данных ЭСКРОУ счетов</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Загрузите Excel-файлы для анализа</p>', unsafe_allow_html=True)

# Боковая панель для настроек
with st.sidebar:
    st.markdown('<h3 style="color: #272727;">Настройки</h3>', unsafe_allow_html=True)
    filter_by_period = st.checkbox("Фильтровать по году и месяцу", value=True, help="Ограничить анализ конкретным периодом")
    
    if filter_by_period:
        col1, col2 = st.columns(2)
        with col1:
            year = st.number_input("Год", min_value=2000, max_value=2100, value=datetime.now().year)
        with col2:
            month = st.selectbox("Месяц", 
                                options=list(range(1, 13)), 
                                format_func=lambda x: datetime(year, x, 1).strftime("%B"), 
                                index=datetime.now().month-1)
    else:
        year, month = None, None

# Область для перетаскивания нескольких Excel-файлов
st.markdown(f'<h3 class="subtitle"> Финансовый анализ ЭСКРОУ счетов ({"Filter by Period" if filter_by_period else "No Filter"})</h3>', unsafe_allow_html=True)
st.markdown("""
<div id="excel_drop_zone">
    Перетащите Excel-файлы или выберите файлы
    <input type="file" id="excel_file_input" accept=".xlsx,.xls" multiple style="display: none;">
</div>
""", unsafe_allow_html=True)

# Загрузчик файлов
uploaded_excel_files = st.file_uploader("Выберите Excel-файлы", type=["xlsx", "xls"], key="excel_uploader", accept_multiple_files=True, label_visibility="collapsed")

# Обработка загруженных файлов
if uploaded_excel_files:
    period_text = f"за {year}-{month:02d}" if filter_by_period else "за весь период"
    st.markdown(f'<p class="subtitle">Обработка данных {period_text}...</p>', unsafe_allow_html=True)
    result_df, error_df = analyze_excel_files(uploaded_excel_files, year, month, filter_by_period)

    if not result_df.empty:
        st.markdown(f'<p class="subtitle">Результаты анализа ({period_text}, положительные суммы):</p>', unsafe_allow_html=True)
        st.dataframe(result_df)

        # Создание файла для скачивания результатов
        output = io.BytesIO()
        result_df.to_excel(output, index=False)
        output.seek(0)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        filename = f"excel_results_{'all_period' if not filter_by_period else f'{year}_{month:02d}'}_{timestamp}.xlsx"
        st.markdown(get_binary_file_downloader_html(output.getvalue(), filename), unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="error-message">Данные {period_text} не обработаны. Проверьте файлы.</div>', unsafe_allow_html=True)

    if not error_df.empty:
        st.markdown('<div class="error-message">Ошибки обработки:</div>', unsafe_allow_html=True)
        st.dataframe(error_df)

        # Создание файла для скачивания ошибок
        error_output = io.BytesIO()
        error_df.to_excel(error_output, index=False)
        error_output.seek(0)
        error_filename = f"errors_{'all_period' if not filter_by_period else f'{year}_{month:02d}'}_{timestamp}.xlsx"
        st.markdown(get_binary_file_downloader_html(error_output.getvalue(), error_filename), unsafe_allow_html=True)
    else:
        st.markdown('<div class="success-message">Обработка завершена успешно.</div>', unsafe_allow_html=True)