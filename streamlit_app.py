# streamlit_app.py
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from io import BytesIO
import datetime
import textwrap

# ---------------------------
# Налаштування сторінки
# ---------------------------
st.set_page_config(page_title="Impact Evaluator", layout="wide")
st.title("📊 Impact Evaluator — Оцінка програми (MEAL / Protection / Legal Aid)")

st.markdown(
    "Цей інструмент поєднує кількісні показники та наративний опис (Word .docx) і генерує "
    "детальну аналітику з рекомендаціями."
)

# ---------------------------
# Sidebar — введення даних
# ---------------------------
st.sidebar.header("Введіть кількісні показники")

program_name = st.sidebar.text_input("Назва програми", value="Нова програма")
period = st.sidebar.text_input("Період (наприклад: 2025 Q1)", value=str(datetime.date.today().year))
location = st.sidebar.text_input("Географія (регіон)", value="—")

# ПОРЯДОК ПОЛІВ: закриті кейси перед загальними (згідно з вимогою)
closed_cases = st.sidebar.number_input("Скільки кейсів успішно закрито?", min_value=0, step=1, value=0)
total_cases = st.sidebar.number_input("Скільки кейсів відкрито загалом?", min_value=1, step=1, value=1)

beneficiaries = st.sidebar.number_input("Скільки бенефіціарів було охоплено?", min_value=0, step=1, value=0)
resources_spent = st.sidebar.number_input("Скільки коштів витрачено (USD)?", min_value=0.0, step=1.0, value=0.0)
staff = st.sidebar.number_input("Скільки співробітників працювало над програмою?", min_value=1, step=1, value=1)
community_activities = st.sidebar.number_input("Скільки заходів community-based protection було проведено?", min_value=0, step=1, value=0)

st.sidebar.markdown("---")
st.sidebar.subheader("Наратив (завантажити Word .docx)")
uploaded_docx = st.sidebar.file_uploader("Завантажити .docx (опис програми, мета, проблеми, контекст)", type=["docx"])

st.sidebar.markdown("---")
run_button = st.sidebar.button("🔍 Провести оцінку")

# ---------------------------
# Допоміжні функції
# ---------------------------
def read_docx(uploaded_file):
    """Читає .docx із Streamlit UploadedFile та повертає текст"""
    if uploaded_file is None:
        return ""
    try:
        # Python-docx підтримує file-like об'єкти
        doc = Document(uploaded_file)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return "\n".join(full_text).strip()
    except Exception as e:
        st.warning(f"Не вдалося прочитати .docx: {e}")
        return ""

def sanitize_positive(x, default=0.0):
    try:
        val = float(x)
        return max(val, 0.0)
    except:
        return default

def build_long_evaluation(narrative_text, metrics, sim_results):
    """Генерує розлогий звіт — поєднуючи наратив та кількісні результати."""
    lines = []
    lines.append(f"Оцінка програми — {metrics['program_name']} ({metrics['period']} — {metrics['location']})")
    lines.append("")
    lines.append("1) Короткий виклад (Executive summary):")
    lines.append(f"- Програма охопила приблизно {metrics['beneficiaries']} бенефіціарів.")
    lines.append(f"- Загальна кількість кейсів: {metrics['total_cases']}, з них закрито: {metrics['closed_cases']} ({metrics['case_closure_rate']*100:.1f}%).")
    lines.append(f"- Витрати: ${metrics['resources_spent']:.2f}. Середня вартість на закритий кейс: ${metrics['cost_per_closed_case']:.2f}.")
    lines.append("")
    # вставка ключових висновків зі симуляцій / кореляцій
    lines.append("2) Аналіз зв'язків (кореляції / сенситивність)")
    if sim_results and sim_results.get("corrs"):
        corrs = sim_results["corrs"]
        lines.append(f"- Кореляція (Pearson r) між витратами на закритий кейс та beneficiaries_per_staff (симуляція): {corrs.get('beneficiaries_per_staff', 0):+.2f}")
        lines.append(f"- Кореляція між витратами на закритий кейс та case_closure_rate (симуляція): {corrs.get('case_closure_rate', 0):+.2f}")
        lines.append(f"- Кореляція між витратами на закритий кейс та community_activities_per_staff (симуляція): {corrs.get('cbp_per_staff', 0):+.2f}")
        lines.append("")
        # interpret correlations
        def interpret_r(r):
            r = float(r)
            if abs(r) < 0.2:
                return "слабкий/відсутній"
            if abs(r) < 0.5:
                return "помірний"
            return "сильний"
        lines.append("Інтерпретація: ")
        lines.append(f"- Зв'язок з beneficiaries_per_staff: {interpret_r(corrs.get('beneficiaries_per_staff',0))}.")
        lines.append(f"- Зв'язок з case_closure_rate: {interpret_r(corrs.get('case_closure_rate',0))}.")
        lines.append(f"- Зв'язок з community_activities_per_staff: {interpret_r(corrs.get('cbp_per_staff',0))}.")
        lines.append("")
    else:
        lines.append("- Немає симуляційних даних для оцінки кореляцій.")
        lines.append("")
    # наративна частина — витягуємо ключові твердження з uploaded narrative (прості heuristics)
    lines.append("3) Наративний аналіз (витягнуто з завантаженого опису):")
    if narrative_text:
        # обрізана версія наративу (перших 800 символів + виявлені ключові слова)
        snippet = narrative_text.strip().replace("\n", " ")
        snippet_short = (snippet[:1000] + "...") if len(snippet) > 1000 else snippet
        lines.append(snippet_short)
        # ключові слова
        keywords = {
            "staff": ["персонал", "staff", "штат", "співробіт", "працівн"],
            "funding": ["грошей", "фінанс", "fund", "бюджет", "витрат"],
            "access": ["доступ", "access", "підтримка", "послуг"],
            "safety": ["безпек", "safety", "насиль", "violence", "protection"]
        }
        found = []
        lower_text = snippet.lower()
        for k, kwlist in keywords.items():
            for kw in kwlist:
                if kw in lower_text:
                    found.append(k)
                    break
        if found:
            lines.append(f"Ключові тематичні вектори в наративі: {', '.join(found)}.")
        else:
            lines.append("У наративі прямо не виявлено чітких згадок про персонал/фінанси/доступ/безпеку (за простим аналізом).")
    else:
        lines.append("- Наратив не завантажено.")
    lines.append("")
    # Рекомендації - обґрунтовані з metrics
    lines.append("4) Рекомендації (з аргументацією):")
    recs = []
    # cost per closed case
    if metrics['cost_per_closed_case'] > 200:
        recs.append(("Оптимізація витрат", 
                     "Вартість на закритий кейс є досить високою. Рекомендовано провести аудит закупівель, "
                     "оптимізувати логістику та розглянути масштабування послуг ( щоб зменшити unit-cost)."))
    else:
        recs.append(("Ефективність витрат", "Вартість на закритий кейс у прийнятних межах; розгляньте реплікацію підходів у інших регіонах."))
    # case closure
    if metrics['case_closure_rate'] < 0.6:
        recs.append(("Покращення кейс-менеджменту", 
                     "Низький рівень закриття кейсів. Перегляньте SOP, флоу обробки, час відгуку та фоллов-ап. Можливо, потрібні тренінги для кейс-воркерів."))
    else:
        recs.append(("Підтримка кейс-менеджменту", "Рівень закриття кейсів є задовільним; документуйте кейс-стаді та best practices."))
    # staff load
    if metrics['beneficiaries_per_staff'] > 80:
        recs.append(("Розвантаження персоналу", "Навантаження на персонал високе — розгляньте найм або автоматизацію повторюваних процесів."))
    else:
        recs.append(("Баланс навантаження", "Навантаження персоналу в межах прийнятних показників."))
    # community activities
    if metrics['community_activities'] < 3:
        recs.append(("Посилення community-based activities", "Низька кількість заходів СВП; розгляньте збільшення активностей для підвищення довіри громади."))
    else:
        recs.append(("Community engagement", "Достатній рівень активностей; фіксуйте вплив в кейс-репортах."))

    # format recommendations
    for title, text in recs:
        lines.append(f"- {title}: {text}")

    lines.append("")
    lines.append("5) Пропоновані наступні кроки:")
    lines.append("- Провести внутрішній аудит по витратах та логістиці (2–4 тижні).")
    lines.append("- Провести ревізію кейс-менеджменту та SOP (1–2 місяці).")
    lines.append("- Розробити план підвищення охоплення через CBP заходи (3–6 місяців).")

    return "\n".join(lines)

def monte_carlo_simulation(metrics, n=300, perturb=0.25):
    """
    Робимо сенситивну симуляцію параметрів навколо введених значень,
    щоб отримати 'штучну' множину точок для оцінки кореляцій.
    Повертаємо DataFrame з змінними та dict кореляцій.
    """
    rng = np.random.default_rng(12345)
    base = metrics.copy()
    samples = []
    for i in range(n):
        # випадкова зміна кожного показника ±perturb
        fac_ben = 1 + rng.normal(0, perturb)
        fac_closed = 1 + rng.normal(0, perturb)
        fac_staff = 1 + rng.normal(0, perturb)
        fac_resources = 1 + rng.normal(0, perturb)
        fac_cbp = 1 + rng.normal(0, perturb)

        beneficiaries = max(0.0, base['beneficiaries'] * fac_ben)
        closed_cases = max(1.0, base['closed_cases'] * fac_closed)
        staff = max(1.0, base['staff'] * fac_staff)
        resources_spent = max(0.0, base['resources_spent'] * fac_resources)
        cbp = max(0.0, base['community_activities'] * fac_cbp)

        cost_per_closed_case = resources_spent / closed_cases
        beneficiaries_per_staff = beneficiaries / staff
        case_closure_rate = closed_cases / max(base['total_cases'], 1.0)  # note: keep denom as original total_cases for realism
        cbp_per_staff = cbp / staff

        samples.append({
            "cost_per_closed_case": cost_per_closed_case,
            "beneficiaries_per_staff": beneficiaries_per_staff,
            "case_closure_rate": case_closure_rate,
            "cbp_per_staff": cbp_per_staff
        })

    df = pd.DataFrame(samples)
    # кореляції (Pearson r)
    corrs = {}
    try:
        corrs["beneficiaries_per_staff"] = np.corrcoef(df["cost_per_closed_case"], df["beneficiaries_per_staff"])[0,1]
        corrs["case_closure_rate"] = np.corrcoef(df["cost_per_closed_case"], df["case_closure_rate"])[0,1]
        corrs["cbp_per_staff"] = np.corrcoef(df["cost_per_closed_case"], df["cbp_per_staff"])[0,1]
    except Exception:
        corrs = {"beneficiaries_per_staff": 0.0, "case_closure_rate": 0.0, "cbp_per_staff": 0.0}

    return df, corrs

# ---------------------------
# MAIN: коли натиснуто кнопку "Провести оцінку"
# ---------------------------
if run_button:
    # sanitize inputs
    beneficiaries = int(sanitize_positive(beneficiaries, 0))
    closed_cases = int(sanitize_positive(closed_cases, 0))
    total_cases = int(sanitize_positive(total_cases, 1))
    resources_spent = sanitize_positive(resources_spent, 0.0)
    staff = int(sanitize_positive(staff, 1))
    community_activities = int(sanitize_positive(community_activities, 0))

    # metrics calculations
    closed_cases = max(closed_cases, 0)
    total_cases = max(total_cases, 1)
    case_closure_rate = closed_cases / total_cases
    cost_per_closed_case = resources_spent / closed_cases if closed_cases > 0 else float('inf')
    beneficiaries_per_staff = beneficiaries / staff if staff > 0 else 0.0

    metrics = {
        "program_name": program_name,
        "period": period,
        "location": location,
        "beneficiaries": beneficiaries,
        "closed_cases": closed_cases,
        "total_cases": total_cases,
        "resources_spent": resources_spent,
        "staff": staff,
        "community_activities": community_activities,
        "case_closure_rate": case_closure_rate,
        "cost_per_closed_case": cost_per_closed_case,
        "beneficiaries_per_staff": beneficiaries_per_staff
    }

    # read docx narrative
    narrative_text = ""
    if uploaded_docx is not None:
        narrative_text = read_docx(uploaded_docx)
        st.info("Наратив успішно завантажено.")
    else:
        st.info("Наратив не завантажено — завантажте .docx у бічній панелі для більш глибокого звіту.")

    # monte carlo simulation to produce scatter for "correlation" visualization
    sim_df, corrs = monte_carlo_simulation(metrics, n=400, perturb=0.20)
    sim_results = {"df": sim_df, "corrs": corrs}

    # ---------------------------
    # Відображення результатів у головній панелі
    # ---------------------------
    st.subheader("📌 Ключові показники")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Бенефіціарів", f"{beneficiaries}")
    col2.metric("Кейсів відкрито (загалом)", f"{total_cases}")
    col3.metric("Кейсів успішно закрито", f"{closed_cases}")
    col4.metric("Case closure rate", f"{case_closure_rate*100:.1f}%")

    st.markdown("---")
    st.subheader("🔎 Детальні числові показники")
    st.write(pd.DataFrame([{
        "program_name": program_name,
        "period": period,
        "location": location,
        "beneficiaries": beneficiaries,
        "total_cases": total_cases,
        "closed_cases": closed_cases,
        "resources_spent": resources_spent,
        "staff": staff,
        "community_activities": community_activities,
        "cost_per_closed_case": round(cost_per_closed_case, 2),
        "beneficiaries_per_staff": round(beneficiaries_per_staff, 2),
        "case_closure_rate (%)": round(case_closure_rate*100, 2)
    }]).T)

    st.markdown("---")
    st.subheader("📈 Аналіз кореляцій (сенситивна симуляція)")

    # Scatter 1: cost_per_closed_case vs beneficiaries_per_staff
    fig1, ax1 = plt.subplots(figsize=(6,3))
    ax1.scatter(sim_df["beneficiaries_per_staff"], sim_df["cost_per_closed_case"], alpha=0.5)
    # trendline
    try:
        z = np.polyfit(sim_df["beneficiaries_per_staff"], sim_df["cost_per_closed_case"], 1)
        p = np.poly1d(z)
        xs = np.linspace(sim_df["beneficiaries_per_staff"].min(), sim_df["beneficiaries_per_staff"].max(), 100)
        ax1.plot(xs, p(xs), color="red", linewidth=1)
    except Exception:
        pass
    ax1.set_xlabel("Beneficiaries per staff")
    ax1.set_ylabel("Cost per closed case (USD)")
    ax1.set_title(f"Витрати на закритий кейс vs beneficiaries_per_staff\nPearson r = {corrs['beneficiaries_per_staff']:+.2f}")
    st.pyplot(fig1)

    # Scatter 2: cost_per_closed_case vs case_closure_rate
    fig2, ax2 = plt.subplots(figsize=(6,3))
    ax2.scatter(sim_df["case_closure_rate"], sim_df["cost_per_closed_case"], alpha=0.5)
    try:
        z2 = np.polyfit(sim_df["case_closure_rate"], sim_df["cost_per_closed_case"], 1)
        p2 = np.poly1d(z2)
        xs2 = np.linspace(sim_df["case_closure_rate"].min(), sim_df["case_closure_rate"].max(), 100)
        ax2.plot(xs2, p2(xs2), color="red", linewidth=1)
    except Exception:
        pass
    ax2.set_xlabel("Case closure rate (fraction)")
    ax2.set_ylabel("Cost per closed case (USD)")
    ax2.set_title(f"Витрати на закритий кейс vs case_closure_rate\nPearson r = {corrs['case_closure_rate']:+.2f}")
    st.pyplot(fig2)

    # Scatter 3: cost_per_closed_case vs cbp_per_staff
    fig3, ax3 = plt.subplots(figsize=(6,3))
    ax3.scatter(sim_df["cbp_per_staff"], sim_df["cost_per_closed_case"], alpha=0.5)
    try:
        z3 = np.polyfit(sim_df["cbp_per_staff"], sim_df["cost_per_closed_case"], 1)
        p3 = np.poly1d(z3)
        xs3 = np.linspace(sim_df["cbp_per_staff"].min(), sim_df["cbp_per_staff"].max(), 100)
        ax3.plot(xs3, p3(xs3), color="red", linewidth=1)
    except Exception:
        pass
    ax3.set_xlabel("Community activities per staff")
    ax3.set_ylabel("Cost per closed case (USD)")
    ax3.set_title(f"Витрати на закритий кейс vs CBP per staff\nPearson r = {corrs['cbp_per_staff']:+.2f}")
    st.pyplot(fig3)

    st.markdown("---")
    st.subheader("🧾 Розгорнутий звіт (на основі наративу та метрик)")
    long_report = build_long_evaluation(narrative_text, metrics, sim_results)
    st.text_area("Evaluation report", value=long_report, height=520)

    st.success("Оцінка завершена — використай звіт і графіки для презентації та планування дій.")
