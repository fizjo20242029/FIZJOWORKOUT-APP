import json
import os
import random
import io
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import streamlit as st
from groq import Groq

# ==============================================================================
# KONFIGURACJA STRONY STREAMLIT
# ==============================================================================
st.set_page_config(page_title="Fizjo Workout Ultimate", page_icon="💪", layout="centered")

# ==============================================================================
# INICJALIZACJA STANU SESJI (PAMIĘĆ APLIKACJI WEBOWEJ)
# ==============================================================================
if 'wylosowany_plan_cache' not in st.session_state:
    st.session_state.wylosowany_plan_cache = []
if 'historia_wiadomosci' not in st.session_state:
    st.session_state.historia_wiadomosci = [
        {"role": "system", "content": "Jesteś wirtualnym asystentem w aplikacji dla fizjoterapeutów i trenerów 'Fizjo Workout Ultimate'. Pomagasz profesjonalnie i zwięźle. Język: polski."}
    ]

# ==============================================================================
# BAZY DANYCH
# ==============================================================================
BAZA_FIZJO = {
    "Oddechowe": [
        {"nazwa": "Oddech przeponowy w leżeniu", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie tyłem, kolana ugięte. Jedna ręka na klatce, druga na brzuchu. Wdech nosem kieruje powietrze do brzucha (brzuch rośnie), wydech ustami.", "czas_min": 2, "parametry": "2-3 minuty", "miesnie": "Przepona, mięśnie międzyżebrowe"},
        {"nazwa": "Oddech metodą 'pursed lips'", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Wdech nosem, a następnie powolny, maksymalnie wydłużony wydech przez lekko przymknięte usta.", "czas_min": 2, "parametry": "2 minuty", "miesnie": "Przepona, mięśnie tłoczni brzusznej"},
        {"nazwa": "Oddech dolnożebrowy z taśmą", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Owiń taśmę elastyczną wokół dolnych żeber. Podczas wdechu staraj się rozepchnąć taśmę na boki.", "czas_min": 2, "parametry": "10-12 powtórzeń", "miesnie": "Mięśnie międzyżebrowe zewnętrzne, przepona"}
    ],
    "Głowa/Szyja": [
        {"nazwa": "Retrakcja szyi (Cofanie brody)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: W pozycji siedzącej cofnij głowę w płaszczyźnie poziomej (zrób 'podwójny podbródek'). Wzrok prosto, trzymaj 3 s.", "czas_min": 2, "parametry": "3 serie x 10 powtórzeń", "miesnie": "Mięśnie głębokie zginacze szyi, mięsień płatowaty głowy"},
        {"nazwa": "Izometryczne parcie w przód", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Przyłóż dłoń do czoła. Naciskaj głową w przód na stawiającą opór dłoń, nie pozwalając na ruch.", "czas_min": 2, "parametry": "3 serie x 5s trzymania", "miesnie": "Mięsień mostkowo-obojczykowo-sutkowy (MOS), zginacze długie"}
    ],
    "Kończyna górna": [
        {"nazwa": "Wznosy ramion w pozycji Y", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie przodem, ramiona pod kątem 45 stopni (litera Y), kciuki w górę. Unoś ramiona, łącząc dolne kąty łopatek.", "czas_min": 3, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Mięsień czworoboczny (część dolna), mięsień równoległoboczny"},
        {"nazwa": "Pompki plus (Scapular)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Pozycja podporu przodem. Bez uginania łokci oddalaj łopatki od siebie i zbliżaj.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Mięsień zębaty przedni, mięsień piersiowy mniejszy"}
    ],
    "Core (Tułów)": [
        {"nazwa": "Plank (Podpór przodem)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Oprzyj się na przedramionach i palcach stóp. Ciało w jednej linii, brzuch i pośladki mocno spięte, miednica neutralnie.", "czas_min": 3, "parametry": "3 serie x 30 sekund", "miesnie": "Mięsień poprzeczny brzucha, mięsień prosty brzucha, pośladkowy wielki"},
        {"nazwa": "Dead Bug (Martwy robak)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie tyłem, ręce pionowo, nogi ugięte 90/90. Opuszczaj jednocześnie przeciwną rękę i nogę tuż nad podłogę, lędźwie w matę.", "czas_min": 3, "parametry": "3 serie x 10 powtórzeń na stronę", "miesnie": "Mięsień poprzeczny brzucha, skośny wewnętrzny"}
    ],
    "Kończyna dolna": [
        {"nazwa": "Przysiad klasyczny (Squat)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Stopy na szerokość barków. Schodź biodrami w dół i w tył, utrzymując kolana w linii stóp oraz zachowując proste plecy.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Mięsień czworogłowy uda, pośladkowy wielki"},
        {"nazwa": "Mostki biodrowe (Glute Bridge)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie tyłem, kolana ugięte, stopy na ziemi. Unoś biodra w górę poprzez mocne spięcie pośladków.", "czas_min": 2, "parametry": "3 serie x 15 powtórzeń", "miesnie": "Mięsień pośladkowy wielki, dwugłowy uda"}
    ]
}

BAZA_SILOWNIA = {
    "Rozgrzewka": [
        {"nazwa": "Monster Walk z mini-bandem", "opis": "INSTRUKCJA: Załóż mini-band nad kolana, przyjmij pozycję półprzysiadu. Wykonuj szerokie kroki w bok.", "czas_min": 2, "parametry": "2 min", "miesnie": "Pośladkowy średni, stabilizatory"}
    ],
    "Klatka piersiowa": [
        {"nazwa": "Wyciskanie sztangi płasko", "opis": "POZYCJA: Połóż się na ławce, stopy stabilnie na podłożu. RUCH: Opuść sztangę do dolnej części klatki, a następnie wyciśnij dynamicznie w górę.", "czas_min": 3, "parametry": "4x8", "miesnie": "Klatka, triceps"}
    ],
    "Plecy": [
        {"nazwa": "Podciąganie na drążku nachwytem", "opis": "POZYCJA: Chwyć drążek szerzej niż barki. RUCH: Podciągnij klatkę piersiową do drążka.", "czas_min": 3, "parametry": "3 serie x Max", "miesnie": "Najszerszy grzbietu, obły większy"}
    ],
    "Ręce": [
        {"nazwa": "Uginanie sztangi łamanej", "opis": "POZYCJA: Stojąc, chwyć sztangę podchwytem. RUCH: Uginaj ramiona w łokciach.", "czas_min": 2, "parametry": "3x12", "miesnie": "Biceps, m. ramienny"}
    ],
    "Nogi": [
        {"nazwa": "Wykroki z hantlami", "opis": "POZYCJA: Stojąc, hantle w dłoniach. RUCH: Wykonaj duży krok w przód, uginając oba kolana.", "czas_min": 3, "parametry": "3x12 na nogę", "miesnie": "Czworogłowe, pośladki"}
    ],
    "Pośladki": [
        {"nazwa": "Hip Thrust (ze sztangą)", "opis": "POZYCJA: Łopatki oparte o ławkę, sztanga na biodrach. RUCH: Wypchnij biodra w górę.", "czas_min": 4, "parametry": "4x10", "miesnie": "Pośladkowy wielki"}
    ],
    "Zakończenie treningu": [
        {"nazwa": "Schładzanie na rowerku", "opis": "POZYCJA: Siedząc na rowerku. RUCH: Spokojne pedałowanie.", "czas_min": 5, "parametry": "5 minut", "miesnie": "Całe ciało, układ krążenia"}
    ]
}

PLIK_WLASNYCH_CWICZEN = "wlasne_cwiczenia.json"

def zaladuj_wlasne_cwiczenia():
    if os.path.exists(PLIK_WLASNYCH_CWICZEN):
        try:
            with open(PLIK_WLASNYCH_CWICZEN, "r", encoding="utf-8") as f:
                dane = json.load(f)
                for kat, lista in dane.get("FIZJO", {}).items():
                    if kat in BAZA_FIZJO: BAZA_FIZJO[kat].extend(lista)
                for kat, lista in dane.get("GYM", {}).items():
                    if kat in BAZA_SILOWNIA: BAZA_SILOWNIA[kat].extend(lista)
        except: pass

zaladuj_wlasne_cwiczenia()
GLOBALNA_BAZA = {**BAZA_FIZJO, **BAZA_SILOWNIA}

# ==============================================================================
# LOGIKA GENERATORÓW
# ==============================================================================
def generuj_plan(profil, budzet, dni):
    plan = []
    realny_czas = 0
    b_fizjo = {k: list(v) for k, v in BAZA_FIZJO.items()}
    b_gym = {k: list(v) for k, v in BAZA_SILOWNIA.items()}

    if profil.startswith("FIZJO:"):
        if "Kompleksowy" in profil:
            if b_fizjo["Oddechowe"]:
                cw_start = random.choice(b_fizjo["Oddechowe"]).copy()
                cw_start["uwagi"] = ""
                plan.append(("Oddechowe (Rozgrzewka)", cw_start))
                realny_czas += cw_start["czas_min"]
                b_fizjo["Oddechowe"].remove(next(c for c in BAZA_FIZJO["Oddechowe"] if c["nazwa"] == cw_start["nazwa"]))

            cw_koniec = None
            if b_fizjo["Oddechowe"]:
                cw_koniec = random.choice(b_fizjo["Oddechowe"]).copy()
                cw_koniec["uwagi"] = ""
            
            czas_koncowy = cw_koniec["czas_min"] if cw_koniec else 0
            lancuch = ["Głowa/Szyja", "Kończyna górna", "Core (Tułów)", "Kończyna dolna"]
            
            puste = 0
            while realny_czas + czas_koncowy < budzet and puste < 10:
                dodano = False
                for kat in lancuch:
                    if b_fizjo.get(kat):
                        cw = random.choice(b_fizjo[kat]).copy()
                        if realny_czas + cw["czas_min"] + czas_koncowy <= budzet:
                            cw["uwagi"] = ""
                            plan.append((kat, cw))
                            realny_czas += cw["czas_min"]
                            b_fizjo[kat].remove(next(c for c in BAZA_FIZJO[kat] if c["nazwa"] == cw["nazwa"]))
                            dodano = True
                if not dodano: puste += 1
                else: puste = 0
                
            if cw_koniec: plan.append(("Oddechowe (Wyciszenie)", cw_koniec))

        else:
            kat = profil.split(" - ")[1].replace("Tylko ", "")
            dostepne = b_fizjo.get(kat, [])
            while realny_czas < budzet and dostepne:
                cw = random.choice(dostepne).copy()
                if realny_czas + cw["czas_min"] <= budzet:
                    cw["uwagi"] = ""
                    plan.append((kat, cw))
                    realny_czas += cw["czas_min"]
                    dostepne.remove(next(c for c in BAZA_FIZJO[kat] if c["nazwa"] == cw["nazwa"]))
                else:
                    break

    elif profil.startswith("GYM:"):
        rozgrzewka = random.choice(b_gym["Rozgrzewka"]).copy()
        rozgrzewka["uwagi"] = ""
        zakonczenie = random.choice(b_gym["Zakończenie treningu"]).copy()
        zakonczenie["uwagi"] = ""

        if "Split" in profil:
            dni_tygodnia = ["Poniedziałek", "Wtorek", "Środa", "Czwartek", "Piątek", "Sobota", "Niedziela"]
            uklady = {
                1: [["Klatka piersiowa", "Plecy", "Nogi", "Ręce", "Pośladki"]],
                2: [["Klatka piersiowa", "Ręce"], ["Plecy", "Nogi", "Pośladki"]],
                3: [["Klatka piersiowa", "Ręce"], ["Nogi", "Pośladki"], ["Plecy"]],
                4: [["Ręce"], ["Klatka piersiowa"], ["Nogi"], ["Plecy"]]
            }
            plan_na_dni = uklady.get(dni, [["Klatka piersiowa"], ["Plecy"], ["Nogi"], ["Ręce"], ["Pośladki"]][:dni])

            for i in range(dni):
                dzien = dni_tygodnia[i]
                partie = plan_na_dni[i]
                plan.append(("NAGŁÓWEK DNIA", {"nazwa": f"{dzien}: {' + '.join([p.upper() for p in partie])}", "opis": "", "czas_min": 0, "parametry": "-", "miesnie": "-", "uwagi": ""}))
                plan.append(("GYM: Rozgrzewka", rozgrzewka.copy()))
                
                for p in partie:
                    dostepne = list(b_gym.get(p, []))
                    dodano = 0
                    while dodano < budzet and dostepne:
                        cw = random.choice(dostepne).copy()
                        cw["uwagi"] = ""
                        plan.append((f"GYM: {p}", cw))
                        dodano += 1
                        dostepne.remove(next(c for c in BAZA_SILOWNIA[p] if c["nazwa"] == cw["nazwa"]))
                plan.append(("GYM: Zakończenie", zakonczenie.copy()))
        else:
            plan.append(("GYM: Rozgrzewka", rozgrzewka))
            if "Ogólnorozwojowy" in profil:
                for p in ["Klatka piersiowa", "Plecy", "Nogi", "Ręce", "Pośladki"]:
                    dostepne = list(b_gym.get(p, []))
                    dodano = 0
                    while dodano < budzet and dostepne:
                        cw = random.choice(dostepne).copy()
                        cw["uwagi"] = ""
                        plan.append((f"GYM: {p}", cw))
                        dodano += 1
                        dostepne.remove(next(c for c in BAZA_SILOWNIA[p] if c["nazwa"] == cw["nazwa"]))
            else:
                kat = profil.split(" - ")[1]
                dostepne = b_gym.get(kat, [])
                dodano = 0
                while dodano < budzet and dostepne:
                    cw = random.choice(dostepne).copy()
                    cw["uwagi"] = ""
                    plan.append((f"GYM: {kat}", cw))
                    dodano += 1
                    dostepne.remove(next(c for c in BAZA_SILOWNIA[kat] if c["nazwa"] == cw["nazwa"]))
            plan.append(("GYM: Zakończenie", zakonczenie))

    st.session_state.wylosowany_plan_cache = plan
    st.session_state.is_gym = profil.startswith("GYM:")

# ==============================================================================
# GENEROWANIE PLIKÓW BEZPOŚREDNIO DO POBRANIA Z PRZEGLĄDARKI
# ==============================================================================
def generuj_docx():
    doc = Document()
    doc.add_heading('ZINTEGROWANA KARTA REHABILITACYJNO-TRENINGOWA', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("_________________________________________________________________").alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for idx, (kat, cw) in enumerate(st.session_state.wylosowany_plan_cache, 1):
        if kat == "NAGŁÓWEK DNIA":
            p = doc.add_paragraph()
            run = p.add_run(f"\n{cw['nazwa']}\n")
            run.bold = True; run.font.size = Pt(12)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue

        p = doc.add_paragraph()
        run = p.add_run(f"{idx}. {kat.upper()}: {cw['nazwa']}\n")
        run.bold = True
        
        p.add_run("DAWKOWANIE: ").bold = True
        p.add_run(f"{cw['parametry']}\n")
        p.add_run("MIESNIE: ").bold = True
        p.add_run(f"{cw['miesnie']}\n")
        p.add_run("OPIS: ").bold = True
        p.add_run(f"{cw['opis']}\n")
        p.add_run("UWAGI: ").bold = True
        p.add_run(f"{cw.get('uwagi', '..........................................................')}\n")
        
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def generuj_excel():
    dane_excel = [["L.p.", "Kategoria", "Nazwa ćwiczenia", "Czas", "Zalecenia", "Sposób wykonania", "Uwagi"]]
    lp = 1
    for kat, cw in st.session_state.wylosowany_plan_cache:
        if kat == "NAGŁÓWEK DNIA":
            dane_excel.append(["", "", "", "", "", "", ""])
            dane_excel.append([cw['nazwa'], "", "", "", "", "", ""])
            lp = 1
        else:
            dane_excel.append([lp, kat, cw['nazwa'], str(cw.get('czas_min','-')), cw['parametry'], cw['opis'], cw.get('uwagi', '')])
            lp += 1
            
    df = pd.DataFrame(dane_excel)
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, header=False, sheet_name='Harmonogram')
    return bio.getvalue()

# ==============================================================================
# UI - INTERFEJS APLIKACJI MOBILNEJ / WEBOWEJ
# ==============================================================================
st.title("Fizjo Workout Ultimate")
st.markdown("Zintegrowane środowisko projektowania programów treningowych.")

# MENU BOCZNE
# MENU BOCZNE
with st.sidebar:
    st.header("🔑 Dostęp do AI")
    user_api_key = st.text_input("Twój klucz API Groq:", type="password", help="Pobierz darmowy klucz ze strony console.groq.com")
    
    # Próba logowania podanym kluczem
    groq_client = None
    if user_api_key:
        try:
            groq_client = Groq(api_key=user_api_key)
            st.success("Klucz API zaakceptowany!")
        except Exception:
            st.error("Nieprawidłowy klucz API.")

    st.divider()
    st.header("⚙️ Konfiguracja Planu")
    
    profil = st.selectbox("Profil Silnika:", [
        "FIZJO: Kompleksowy (Wszystkie partie)", 
        "FIZJO: Ukierunkowany - Tylko Oddechowe",
        "FIZJO: Ukierunkowany - Tylko Głowa/Szyja",
        "FIZJO: Ukierunkowany - Tylko Kończyna górna",
        "FIZJO: Ukierunkowany - Tylko Core (Tułów)",
        "FIZJO: Ukierunkowany - Tylko Kończyna dolna",
        "GYM: Ogólnorozwojowy (FBW - Całe Ciało)",
        "GYM: Automatyczny Split (Dni Tygodnia)",
        "GYM: Ukierunkowany - Klatka piersiowa",
        "GYM: Ukierunkowany - Ręce",
        "GYM: Ukierunkowany - Plecy",
        "GYM: Ukierunkowany - Nogi",
        "GYM: Ukierunkowany - Pośladki"
    ])
    
    is_gym = profil.startswith("GYM:")
    label_param = "Ilość ćw. NA PARTIĘ:" if is_gym else "Budżet czasu (min):"
    domyslna_wartosc = 4 if is_gym else 45
    
    budzet = st.number_input(label_param, min_value=1, max_value=120, value=domyslna_wartosc)
    dni = st.number_input("Liczba dni (dla Split):", min_value=1, max_value=7, value=4)
    
    if st.button("⚡ GENERUJ AUTOMAT", use_container_width=True, type="primary"):
        generuj_plan(profil, budzet, dni)
        st.rerun()
        
    st.divider()
    
    if st.session_state.wylosowany_plan_cache:
        st.success("Plan gotowy do pobrania!")
        st.download_button(
            label="💾 POBIERZ DOCX",
            data=generuj_docx(),
            file_name="Plan_Treningowy.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        st.download_button(
            label="📊 POBIERZ EXCEL",
            data=generuj_excel(),
            file_name="Plan_Treningowy.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ZAKŁADKI GŁÓWNE
tab1, tab2, tab3 = st.tabs(["📝 Twój Plan", "➕ Kreator", "✨ Asystent AI Groq"])

# ZAKŁADKA 1: WYGENEROWANY PLAN
with tab1:
    if not st.session_state.wylosowany_plan_cache:
        st.info("👈 Użyj panelu bocznego, aby wygenerować swój pierwszy plan.")
    else:
        for idx, (kat, cw) in enumerate(st.session_state.wylosowany_plan_cache):
            if kat == "NAGŁÓWEK DNIA":
                st.markdown(f"### 📅 {cw['nazwa']}")
                continue
                
            with st.expander(f"{idx+1}. {cw['nazwa']} ({kat})", expanded=True):
                col1, col2 = st.columns([3, 1])
                with col1:
                    nowe_parametry = st.text_input("Zalecenie", cw['parametry'], key=f"p_{idx}")
                    nowe_uwagi = st.text_input("Uwagi dla pacjenta", cw.get('uwagi', ''), key=f"u_{idx}")
                    
                    st.session_state.wylosowany_plan_cache[idx][1]['parametry'] = nowe_parametry
                    st.session_state.wylosowany_plan_cache[idx][1]['uwagi'] = nowe_uwagi
                    
                    st.caption(f"**Anatomia:** {cw['miesnie']}")
                    st.write(cw['opis'])
                with col2:
                    if st.button("Usuń", key=f"del_{idx}", type="secondary"):
                        st.session_state.wylosowany_plan_cache.pop(idx)
                        st.rerun()

# ZAKŁADKA 2: KREATOR MANUALNY
with tab2:
    st.subheader("Baza wszystkich ćwiczeń")
    wybrana_kategoria = st.selectbox("Wybierz kategorię do przeglądu:", list(GLOBALNA_BAZA.keys()))
    
    for cw in GLOBALNA_BAZA[wybrana_kategoria]:
        col1, col2 = st.columns([4, 1])
        with col1:
            st.write(f"**{cw['nazwa']}** [{cw['miesnie']}]")
        with col2:
            if st.button("Dodaj do planu", key=f"add_{cw['nazwa']}"):
                cw_kopia = cw.copy()
                cw_kopia["uwagi"] = ""
                etykieta = f"GYM: {wybrana_kategoria}" if wybrana_kategoria in BAZA_SILOWNIA else wybrana_kategoria
                st.session_state.wylosowany_plan_cache.append((etykieta, cw_kopia))
                st.toast(f"Dodano: {cw['nazwa']}")

# ZAKŁADKA 3: CZAT AI GROQ
# ZAKŁADKA 3: CZAT AI GROQ
with tab3:
    st.subheader("Wirtualny Konsultant Treningowy (Llama 3)")
    
    if not groq_client:
        st.info("👈 Aby rozpocząć czat ze sztuczną inteligencją, wklej swój darmowy klucz API w panelu bocznym. Klucz możesz wygenerować bezpłatnie na stronie console.groq.com.")
    else:
        for msg in st.session_state.historia_wiadomosci:
            if msg["role"] != "system":
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])
        
        if prompt := st.chat_input("O co chcesz zapytać sztuczną inteligencję?"):
            with st.chat_message("user"):
                st.markdown(prompt)
            
            st.session_state.historia_wiadomosci.append({"role": "user", "content": prompt})
            
            with st.chat_message("assistant"):
                message_placeholder = st.empty()
                try:
                    completion = groq_client.chat.completions.create(
                        model="llama-3.3-70b-versatile",
                        messages=st.session_state.historia_wiadomosci,
                        temperature=0.7,
                        max_tokens=1024
                    )
                    odpowiedz = completion.choices[0].message.content
                    message_placeholder.markdown(odpowiedz)
                    
                    st.session_state.historia_wiadomosci.append({"role": "assistant", "content": odpowiedz})
                except Exception as e:
                    st.error(f"Błąd połączenia: {e}")
