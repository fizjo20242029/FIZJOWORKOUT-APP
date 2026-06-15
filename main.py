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

if 'wylosowany_plan_cache' not in st.session_state:
    st.session_state.wylosowany_plan_cache = []
if 'historia_wiadomosci' not in st.session_state:
    st.session_state.historia_wiadomosci = [
        {"role": "system", "content": "Jesteś wirtualnym asystentem w aplikacji dla fizjoterapeutów i trenerów 'Fizjo Workout Ultimate'."}
    ]

# ==============================================================================
# BAZA FIZJOTERAPEUTYCZNA
# ==============================================================================
BAZA_FIZJO = {
    "Oddechowe": [
        {"nazwa": "Oddech przeponowy w leżeniu", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie tyłem, kolana ugięte. Jedna ręka na klatce, druga na brzuchu. Wdech nosem kieruje powietrze do brzucha (brzuch rośnie), wydech ustami.", "czas_min": 2, "parametry": "2-3 minuty", "miesnie": "Przepona, mięśnie międzyżebrowe"},
        {"nazwa": "Oddech metodą 'pursed lips'", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Wdech nosem, a następnie powolny, maksymalnie wydłużony wydech przez lekko przymknięte usta (jak przy delikatnym dmuchaniu na świeczkę).", "czas_min": 2, "parametry": "2 minuty", "miesnie": "Przepona, mięśnie tłoczni brzusznej"},
        {"nazwa": "Oddech dolnożebrowy z taśmą", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Owiń taśmę elastyczną wokół dolnych żeber. Podczas wdechu staraj się rozepchnąć taśmę na boki, podczas wydechu taśma lekko uciska żebra.", "czas_min": 2, "parametry": "10-12 powtórzeń", "miesnie": "Mięśnie międzyżebrowe zewnętrzne, przepona"},
        {"nazwa": "Oddech asymetryczny na boku", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie na boku nieobjętym procesem chorobowym. Podczas wdechu kieruj powietrze do boku leżącego wyżej, rozciągając przestrzenie międzyżebrowe.", "czas_min": 2, "parametry": "5-8 powtórzeń na stronę", "miesnie": "Mięsień czworoboczny lędźwi, międzyżebrowe"},
        {"nazwa": "Pozycja krzesełkowa z oporem", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Siedząc, tułów lekko pochylony w przód, oparcie na przedramionach. Wydłużony wydech z oporem (np. dmuchanie przez rurkę do wody).", "czas_min": 2, "parametry": "2 minuty", "miesnie": "Mięśnie pomocnicze wydechowe, przepona"},
        {"nazwa": "Wdech ze wznosem ramion", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Dynamiczny wdech nosem połączony z uniesieniem ramion przodem w górę, spokojny wydech ustami z opuszczaniem ramion luźno w dół.", "czas_min": 2, "parametry": "12 powtórzeń", "miesnie": "Mięsień piersiowy większy, zębaty przedni"},
        {"nazwa": "Oddech pudełkowy (Box Breathing)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Wdech przez 4 sekundy, zatrzymanie powietrza na 4 sekundy, wydech przez 4 sekundy, zatrzymanie na bezdechu na 4 sekundy.", "czas_min": 2, "parametry": "4 pełne cykle", "miesnie": "Przepona, stabilizatory głębokie tułowia"},
        {"nazwa": "Mobilizacja klatki w rotacji", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Siad klęczny, jedna ręka za głową. Wykonuj powolny skręt tułowia w bok z głębokim wdechem w maksymalnym zakresie, powrót z wydechem.", "czas_min": 2, "parametry": "8-10 powtórzeń", "miesnie": "Mięśnie skośne brzucha, prostownik grzbietu piersiowy"},
        {"nazwa": "Oddech jednostronny w pozycji Sfinksa", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie przodem na przedramionach. Głęboki wdech z intencją skierowania powietrza do tylnych i dolnych partii płuc.", "czas_min": 2, "parametry": "2 minuty", "miesnie": "Przepona, prostownik grzbietu"},
        {"nazwa": "Głęboki wdech w pozycji embrionalnej", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Ukłon japoński. Głębokie wdechy rozszerzające tylną ścianę klatki piersiowej.", "czas_min": 3, "parametry": "3 minuty", "miesnie": "Mięśnie międzyżebrowe tylne, przepona"}
    ],
    "Głowa/Szyja": [
        {"nazwa": "Retrakcja szyi (Cofanie brody)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: W pozycji siedzącej cofnij głowę w płaszczyźnie poziomej (zrób 'podwójny podbródek'). Wzrok prosto, trzymaj 3 s.", "czas_min": 2, "parametry": "3 serie x 10 powtórzeń", "miesnie": "Mięśnie głębokie zginacze szyi, mięsień płatowaty głowy"},
        {"nazwa": "Izometryczne parcie w przód", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Przyłóż dłoń do czoła. Naciskaj głową w przód na stawiającą opór dłoń, nie pozwalając na ruch. Trzymaj przez 5 sekund.", "czas_min": 2, "parametry": "3 serie x 5s trzymania", "miesnie": "MOS, zginacze długie"},
        {"nazwa": "Izometryczne parcie w bok", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Przyłóż dłoń nad uchem. Naciskaj głową w bok przeciwko oporowi ręki. Trzymaj 5 sekund, zmień strony.", "czas_min": 2, "parametry": "3 serie x 5s na stronę", "miesnie": "Mięśnie pochyłe, mięsień płatowaty szyi"},
        {"nazwa": "Izometryczne parcie w tył", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Spleć dłonie z tyłu głowy. Naciskaj tyłem głowy w dłonie, utrzymując wzrok prosto. Utrzymaj 5 sekund.", "czas_min": 2, "parametry": "3 serie x 5s trzymania", "miesnie": "Mięsień podpotyliczny, prostownik grzbietu szyjny"},
        {"nazwa": "Rozciąganie trapeziusa", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Siedząc, opuść jedno ramię w dół. Drugą ręką delikatnie przyciągaj głowę DO przeciwnego barku.", "czas_min": 2, "parametry": "2 serie x 30 sekund", "miesnie": "Czworoboczny (góra), dźwigacz łopatki"},
        {"nazwa": "Rozciąganie dźwigacza łopatki", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Skręć głowę o 45 stopni w bok, a następnie skieruj brodę w dół do pachy. Delikatnie wspomóż ruch ręką.", "czas_min": 2, "parametry": "2 serie x 30 sekund na stronę", "miesnie": "Dźwigacz łopatki"},
        {"nazwa": "Otwieranie ust z oporem", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Podeprzyj żuchwę od dołu palcami. Otwieraj usta powoli, pokonując delikatny opór.", "czas_min": 2, "parametry": "3 serie x 10 powtórzeń", "miesnie": "Mięśnie nadgnykowe"},
        {"nazwa": "Ruchy sakadyczne oczu", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Trzymaj głowę nieruchomo. Przenoś wzrok szybko między dwoma punktami.", "czas_min": 1, "parametry": "1 minuta", "miesnie": "Mięśnie gałkoruchowe"},
        {"nazwa": "Automobilizacja z wałkiem", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie tyłem, twardy wałek pod podstawą czaszki. Wykonuj małe ruchy potakiwania.", "czas_min": 3, "parametry": "3 minuty", "miesnie": "Podpotyliczne"},
        {"nazwa": "Retrakcja z rotacją w leżeniu", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie przodem, czoło nad podłogą. Wykonuj retrakcję szyi, a następnie delikatny obrót głowy.", "czas_min": 2, "parametry": "10 powtórzeń na stronę", "miesnie": "Długi szyi"}
    ],
    "Kończyna górna": [
        {"nazwa": "Wznosy ramion w pozycji Y", "opis": "INSTRUKCJA: Leżenie przodem, ramiona pod kątem 45 stopni (litera Y), kciuki w górę. Unoś ramiona.", "czas_min": 3, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Czworoboczny dolny"},
        {"nazwa": "Rotacja zewnętrzna z gumą", "opis": "INSTRUKCJA: Łokcie ugięte 90 stopni przy bokach. Rozciągaj gumę na boki poprzez rotację zewnętrzną w barkach.", "czas_min": 2, "parametry": "3 serie x 15 powtórzeń", "miesnie": "Stożek rotatorów"},
        {"nazwa": "Pompki plus (Scapular)", "opis": "INSTRUKCJA: W podporze przodem wypychaj łopatki w przód bez uginania łokci.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Zębaty przedni"},
        {"nazwa": "Rozciąganie zginaczy nadgarstka", "opis": "INSTRUKCJA: Wyprostuj rękę, dłoń w górę. Drugą ręką chwyć palce i przyciągaj do siebie.", "czas_min": 2, "parametry": "2 serie x 30 sekund", "miesnie": "Zginacze nadgarstka"},
        {"nazwa": "Rozciąganie prostowników nadgarstka", "opis": "INSTRUKCJA: Wyprostuj rękę, dłoń w dół. Dociśnij dłoń w dół, zginając nadgarstek.", "czas_min": 2, "parametry": "2 serie x 30 sekund", "miesnie": "Prostowniki nadgarstka"},
        {"nazwa": "Odwodzenie hantli w opadzie", "opis": "INSTRUKCJA: Opad tułowia. Unoś ramiona na boki do wysokości tułowia.", "czas_min": 3, "parametry": "3 serie x 10 powtórzeń", "miesnie": "Naramienny tylny"},
        {"nazwa": "Ślizganie ramion po ścianie", "opis": "INSTRUKCJA: Oprzyj plecy o ścianę. Ręce ugięte 90 stopni. Przesuwaj ramiona w górę i w dół.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Czworoboczny dolny"},
        {"nazwa": "Uginanie przedramion z supinacją", "opis": "INSTRUKCJA: W staniu, hantle chwytem neutralnym. Uginaj łokcie rotując dłonie.", "czas_min": 2, "parametry": "3 serie x 15 powtórzeń", "miesnie": "Biceps"},
        {"nazwa": "Prostowanie przedramion w opadzie", "opis": "INSTRUKCJA: Opad tułowia, łokcie 90 stopni. Prostuj ramię w tył.", "czas_min": 3, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Triceps"},
        {"nazwa": "Spacer farmera z kettlebell", "opis": "INSTRUKCJA: Chwyć odważniki. Idź powolnym krokiem ze stabilnymi barkami.", "czas_min": 3, "parametry": "3 serie x 45 sekund", "miesnie": "Przedramiona, góra pleców"}
    ],
    "Core (Tułów)": [
        {"nazwa": "Plank (Podpór przodem)", "opis": "INSTRUKCJA: Oprzyj się na przedramionach i palcach stóp. Ciało w linii prostej.", "czas_min": 3, "parametry": "3 serie x 30 sekund", "miesnie": "Core"},
        {"nazwa": "Side Plank (Podpór bokiem)", "opis": "INSTRUKCJA: Leżenie na boku, podparcie na przedramieniu. Unieś biodra.", "czas_min": 3, "parametry": "2 serie x 20 sekund na stronę", "miesnie": "Skośne brzucha"},
        {"nazwa": "Dead Bug (Martwy robak)", "opis": "INSTRUKCJA: Leżenie tyłem. Opuszczaj przeciwną rękę i nogę, lędźwie w matę.", "czas_min": 3, "parametry": "3 serie x 10 powtórzeń na stronę", "miesnie": "Poprzeczny brzucha"},
        {"nazwa": "Bird Dog (Pies-ptak)", "opis": "INSTRUKCJA: Klęk podparty. Wyciągnij prawą rękę w przód i lewą nogę w tył.", "czas_min": 3, "parametry": "3 serie x 10 powtórzeń na stronę", "miesnie": "Prostownik grzbietu"},
        {"nazwa": "Hollow Body (Pozycja kołyski)", "opis": "INSTRUKCJA: Leżenie tyłem. Oderwij łopatki i proste nogi od podłogi.", "czas_min": 2, "parametry": "3 serie x 20 sekund", "miesnie": "Prosty brzucha"},
        {"nazwa": "Bear Crawl Hold (Niedźwiedź)", "opis": "INSTRUKCJA: Klęk podparty. Unieś kolana 2 cm nad ziemię.", "czas_min": 2, "parametry": "3 serie x 20 sekund", "miesnie": "Poprzeczny brzucha"},
        {"nazwa": "Niedźwiedź dynamiczny", "opis": "INSTRUKCJA: Kroki w podporze niedźwiedzia.", "czas_min": 3, "parametry": "3 serie x 30 sekund", "miesnie": "Core"},
        {"nazwa": "Opuszczanie obustronne nóg", "opis": "INSTRUKCJA: Leżenie tyłem. Powoli opuszczaj obie proste nogi.", "czas_min": 2, "parametry": "3 serie x 10 powtórzeń", "miesnie": "Dół brzucha"},
        {"nazwa": "Skośne spięcia brzucha", "opis": "INSTRUKCJA: Leżenie tyłem. Kieruj ramię w stronę przeciwnego kolana.", "czas_min": 2, "parametry": "3 serie x 15 powtórzeń", "miesnie": "Skośne brzucha"},
        {"nazwa": "Russian Twist", "opis": "INSTRUKCJA: Siad z uniesionymi stopami. Rotuj klatkę piersiową.", "czas_min": 2, "parametry": "3 serie x 20 skrętów", "miesnie": "Skośne brzucha"}
    ],
    "Kończyna dolna": [
        {"nazwa": "Przysiad klasyczny (Squat)", "opis": "INSTRUKCJA: Stopy na szerokość barków. Schodź biodrami w dół i w tył.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Czworogłowy uda, pośladki"},
        {"nazwa": "Mostki biodrowe (Glute Bridge)", "opis": "INSTRUKCJA: Leżenie tyłem, kolana ugięte. Unoś biodra w górę.", "czas_min": 2, "parametry": "3 serie x 15 powtórzeń", "miesnie": "Pośladkowy wielki"},
        {"nazwa": "Wznosy na palce stojąc", "opis": "INSTRUKCJA: Unoś pięty maksymalnie w górę, przechodząc na palce.", "czas_min": 2, "parametry": "3 serie x 20 powtórzeń", "miesnie": "Łydki"},
        {"nazwa": "Zakroki (Reverse Lunges)", "opis": "INSTRUKCJA: Duży krok w tył, opuszczając biodra w dół.", "czas_min": 3, "parametry": "3 serie x 10 powtórzeń na nogę", "miesnie": "Czworogłowy uda"},
        {"nazwa": "Odwodzenie nogi z gumą", "opis": "INSTRUKCJA: Guma nad kostkami. Odprowadzaj nogę w bok.", "czas_min": 2, "parametry": "3 serie x 12 na stronę", "miesnie": "Pośladkowy średni"},
        {"nazwa": "Przysiad bułgarski", "opis": "INSTRUKCJA: Jedna stopa oparta z tyłu. Przysiad na nodze wykrocznej.", "czas_min": 3, "parametry": "3 serie x 8 na nogę", "miesnie": "Czworogłowy uda"},
        {"nazwa": "Clamshell (Muszelka)", "opis": "INSTRUKCJA: Leżenie na boku, kolana ugięte. Unieś górne kolano.", "czas_min": 2, "parametry": "3 serie x 15 na stronę", "miesnie": "Pośladkowy średni"},
        {"nazwa": "Unoszenie prostej nogi (SLR)", "opis": "INSTRUKCJA: Unoś prostą nogę do wysokości drugiego kolana.", "czas_min": 2, "parametry": "3 serie x 12 na nogę", "miesnie": "Prosty uda"},
        {"nazwa": "Martwy ciąg na jednej nodze", "opis": "INSTRUKCJA: Stojąc na jednej nodze, wykonaj skłon w przód.", "czas_min": 3, "parametry": "3 serie x 8 na stronę", "miesnie": "Dwugłowy uda"},
        {"nazwa": "Krzesełko przy ścianie", "opis": "INSTRUKCJA: Oprzyj plecy o ścianę, zejdź biodrami do kąta 90 stopni. Trzymaj nieruchomo.", "czas_min": 2, "parametry": "3x30 sekund", "miesnie": "Czworogłowy uda, stabilizacja kolana"}
    ]
}

# ==============================================================================
# BAZA TRENINGOWA NA SIŁOWNI
# ==============================================================================
BAZA_SILOWNIA = {
    "Rozgrzewka": [
        {"nazwa": "Monster Walk z mini-bandem", "opis": "INSTRUKCJA: Wykonuj szerokie kroki w bok, pilnując kolan na zewnątrz.", "czas_min": 2, "parametry": "2 min", "miesnie": "Pośladkowy średni"},
        {"nazwa": "Face Pulls na wyciągu", "opis": "INSTRUKCJA: Przyciągaj linę do twarzy, rozchylając dłonie.", "czas_min": 2, "parametry": "3x15", "miesnie": "Tył barków"},
        {"nazwa": "Skakanka", "opis": "INSTRUKCJA: Skoki przez skakankę w stałym tempie.", "czas_min": 3, "parametry": "3 min", "miesnie": "Łydki, serce"},
        {"nazwa": "Wiosłowanie taśmą", "opis": "INSTRUKCJA: Zaczep taśmę o słup i przyciągaj końce do bioder.", "czas_min": 2, "parametry": "20 razy", "miesnie": "Plecy"}
    ],
    "Klatka piersiowa": [
        {"nazwa": "Wyciskanie sztangi płasko", "opis": "POZYCJA: Opuść sztangę do dolnej części klatki, a następnie wyciśnij.", "czas_min": 3, "parametry": "4x8", "miesnie": "Klatka, triceps"},
        {"nazwa": "Wyciskanie hantli skos dodatni", "opis": "POZYCJA: Ławka 30 stopni. Wyciskaj hantle nad klatkę.", "czas_min": 3, "parametry": "4x10", "miesnie": "Góra klatki"},
        {"nazwa": "Rozpiętki z hantlami", "opis": "POZYCJA: Leżąc płasko. Otwieraj ramiona szeroko na boki.", "czas_min": 2, "parametry": "3x12", "miesnie": "Klatka"},
        {"nazwa": "Brama (Cable Crossover)", "opis": "POZYCJA: Ściągaj uchwyty wyciągu w dół i do środka.", "czas_min": 3, "parametry": "3x15", "miesnie": "Klatka"}
    ],
    "Plecy": [
        {"nazwa": "Podciąganie na drążku nachwytem", "opis": "POZYCJA: Podciągnij klatkę piersiową do drążka.", "czas_min": 3, "parametry": "3xMax", "miesnie": "Najszerszy grzbietu"},
        {"nazwa": "Wiosłowanie sztangą w opadzie", "opis": "POZYCJA: Przyciągaj sztangę do brzucha, ściągając łopatki.", "czas_min": 3, "parametry": "4x10", "miesnie": "Plecy"},
        {"nazwa": "Ściąganie drążka wyciągu górnego", "opis": "POZYCJA: Przyciągaj drążek do klatki.", "czas_min": 2, "parametry": "3x12", "miesnie": "Najszerszy grzbietu"},
        {"nazwa": "Martwy ciąg klasyczny", "opis": "POZYCJA: Podnieś sztangę, prostując biodra.", "czas_min": 4, "parametry": "4x6", "miesnie": "Plecy, prostowniki"}
    ],
    "Ręce": [
        {"nazwa": "Uginanie sztangi łamanej", "opis": "POZYCJA: Uginaj ramiona w łokciach ze sztangą.", "czas_min": 2, "parametry": "3x12", "miesnie": "Biceps"},
        {"nazwa": "Wyciskanie francuskie", "opis": "POZYCJA: Leżąc na ławce. Opuszczaj sztangę do czoła.", "czas_min": 2, "parametry": "3x10", "miesnie": "Triceps"},
        {"nazwa": "Uginanie hantli z supinacją", "opis": "POZYCJA: Uginaj ramię, obracając nadgarstek.", "czas_min": 2, "parametry": "3x12", "miesnie": "Biceps"},
        {"nazwa": "Prostowanie linek na wyciągu", "opis": "POZYCJA: Prostuj ramiona w dół.", "czas_min": 2, "parametry": "3x15", "miesnie": "Triceps"}
    ],
    "Nogi": [
        {"nazwa": "Przysiad ze sztangą na plecach", "opis": "POZYCJA: Sztanga na karku. Wypchnij biodra w tył.", "czas_min": 4, "parametry": "4x8", "miesnie": "Czworogłowe"},
        {"nazwa": "Wypychanie na suwnicy", "opis": "POZYCJA: Wypychaj platformę nogami.", "czas_min": 3, "parametry": "4x12", "miesnie": "Nogi"},
        {"nazwa": "Martwy ciąg na prostych nogach", "opis": "POZYCJA: Wypychaj biodra w tył z prostymi nogami.", "czas_min": 3, "parametry": "4x10", "miesnie": "Dwugłowe"},
        {"nazwa": "Wykroki z hantlami", "opis": "POZYCJA: Wykonaj duży krok w przód z hantlami.", "czas_min": 3, "parametry": "3x12 na nogę", "miesnie": "Czworogłowe"}
    ],
    "Pośladki": [
        {"nazwa": "Hip Thrust (ze sztangą)", "opis": "POZYCJA: Łopatki oparte o ławkę. Wypchnij biodra w górę.", "czas_min": 4, "parametry": "4x10", "miesnie": "Pośladkowy wielki"},
        {"nazwa": "Odwodzenie nogi na wyciągu", "opis": "POZYCJA: Odprowadzaj prostą nogę w tył.", "czas_min": 2, "parametry": "3x15", "miesnie": "Pośladkowy"},
        {"nazwa": "Glute Bridge z gumą", "opis": "POZYCJA: Leżenie tyłem. Unoś biodra, rozpychając gumę.", "czas_min": 2, "parametry": "3x20", "miesnie": "Pośladki"},
        {"nazwa": "Wykroki boczne", "opis": "POZYCJA: Przenieś ciężar na nogę w boku.", "czas_min": 3, "parametry": "3x12", "miesnie": "Pośladki"}
    ],
    "Zakończenie treningu": [
        {"nazwa": "Schładzanie na rowerku", "opis": "POZYCJA: Spokojne pedałowanie.", "czas_min": 5, "parametry": "5 minut", "miesnie": "Całe ciało"},
        {"nazwa": "Rozciąganie na macie", "opis": "POZYCJA: Klasyczne rozciąganie taśm.", "czas_min": 5, "parametry": "5 minut", "miesnie": "Rozciąganie"}
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

# --- NOWOŚĆ: AUTOMATYCZNE SORTOWANIE ALFABETYCZNE BAZ (PKT 7) ---
for baza in [BAZA_FIZJO, BAZA_SILOWNIA, GLOBALNA_BAZA]:
    for kat in baza:
        baza[kat] = sorted(baza[kat], key=lambda x: x['nazwa'].lower())

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
            # --- NOWOŚĆ: NAPRAWA DLA 5, 6 i 7 DNI SPLITU (PKT 5) ---
            uklady = {
                1: [["Klatka piersiowa", "Plecy", "Nogi", "Ręce", "Pośladki"]],
                2: [["Klatka piersiowa", "Ręce"], ["Plecy", "Nogi", "Pośladki"]],
                3: [["Klatka piersiowa", "Ręce"], ["Nogi", "Pośladki"], ["Plecy"]],
                4: [["Ręce"], ["Klatka piersiowa"], ["Nogi"], ["Plecy"]],
                5: [["Klatka piersiowa"], ["Plecy"], ["Nogi"], ["Ręce"], ["Pośladki"]],
                6: [["Klatka piersiowa"], ["Plecy"], ["Nogi"], ["Ręce"], ["Pośladki"], ["Klatka piersiowa", "Plecy"]],
                7: [["Klatka piersiowa"], ["Plecy"], ["Nogi"], ["Ręce"], ["Pośladki"], ["Klatka piersiowa"], ["Plecy", "Nogi"]]
            }
            plan_na_dni = uklady.get(dni, uklady[7][:dni])

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
# EKSPORTY DO PLIKÓW
# ==============================================================================
def generuj_docx():
    doc = Document()
    doc.add_heading('ZINTEGROWANA KARTA REHABILITACYJNO-TRENINGOWA', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    for idx, (kat, cw) in enumerate(st.session_state.wylosowany_plan_cache, 1):
        if kat == "NAGŁÓWEK DNIA":
            p = doc.add_paragraph()
            run = p.add_run(f"\n{cw['nazwa']}\n")
            run.bold = True; run.font.size = Pt(12)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue
        p = doc.add_paragraph()
        p.add_run(f"{idx}. {kat.upper()}: {cw['nazwa']}\n").bold = True
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

def generuj_excel(liczba_dni):
    if not st.session_state.wylosowany_plan_cache: return None
    dane_excel = [["Imię i Nazwisko:", "", "", "", "", "", ""], ["Płeć:", "", "", "", "", "", ""], ["", "", "", "", "", "", ""]]
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

with st.sidebar:
    st.header("🔑 Dostęp do AI")
    user_api_key = st.text_input("Twój klucz API Groq:", type="password")
    groq_client = Groq(api_key=user_api_key) if user_api_key else None

    st.divider()

    st.header("⚙️ Konfiguracja")
    
    # --- NOWOŚĆ: RADIO ZAMIAST SELECTBOX DLA TELEFONÓW (PKT 1 i 2) ---
    profil = st.radio("Profil Silnika:", [
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
    budzet = st.number_input("Ilość ćw. NA PARTIĘ:" if is_gym else "Budżet czasu (min):", min_value=1, max_value=120, value=4 if is_gym else 45)
    dni = st.number_input("Liczba dni (dla Split):", min_value=1, max_value=7, value=4)
    
    if st.button("⚡ GENERUJ AUTOMAT", use_container_width=True, type="primary"):
        generuj_plan(profil, budzet, dni)
        st.rerun()
        
    if st.button("❌ CZYŚĆ EKRAN (RESET)", use_container_width=True):
        st.session_state.wylosowany_plan_cache = []
        st.rerun()
        
    st.divider()
    if st.session_state.wylosowany_plan_cache:
        st.success("Plan gotowy!")
        st.download_button("💾 POBIERZ DOCX", generuj_docx(), "Plan.docx", use_container_width=True)

# ZAKŁADKI GŁÓWNE
tab1, tab2, tab3, tab4 = st.tabs(["📝 Twój Plan", "➕ Kreator", "✨ Asystent AI", "⚙️ Baza Ćwiczeń"])

# ZAKŁADKA 1: WYGENEROWANY PLAN
with tab1:
    if not st.session_state.wylosowany_plan_cache:
        st.info("👈 Wygeneruj plan w panelu bocznym.")
    else:
        for idx, (kat, cw) in enumerate(st.session_state.wylosowany_plan_cache):
            if kat == "NAGŁÓWEK DNIA":
                st.markdown(f"### 📅 {cw['nazwa']}")
                if st.button("Usuń dzień", key=f"del_day_{idx}", type="primary"):
                    st.session_state.wylosowany_plan_cache.pop(idx)
                    st.rerun()
                continue
                
            with st.expander(f"{idx+1}. {cw['nazwa']} ({kat})", expanded=True):
                col1, col2 = st.columns([3, 2])
                with col1:
                    nowe_parametry = st.text_input("Zalecenie", cw['parametry'], key=f"p_{idx}")
                    nowe_uwagi = st.text_input("Uwagi", cw.get('uwagi', ''), key=f"u_{idx}")
                    st.session_state.wylosowany_plan_cache[idx][1]['parametry'] = nowe_parametry
                    st.session_state.wylosowany_plan_cache[idx][1]['uwagi'] = nowe_uwagi
                    st.caption(f"**Anatomia:** {cw['miesnie']}")
                with col2:
                    c_up, c_down, c_del = st.columns(3)
                    if idx > 0 and c_up.button("⬆️", key=f"up_{idx}"):
                        st.session_state.wylosowany_plan_cache[idx], st.session_state.wylosowany_plan_cache[idx-1] = st.session_state.wylosowany_plan_cache[idx-1], st.session_state.wylosowany_plan_cache[idx]
                        st.rerun()
                    if idx < len(st.session_state.wylosowany_plan_cache) - 1 and c_down.button("⬇️", key=f"down_{idx}"):
                        st.session_state.wylosowany_plan_cache[idx], st.session_state.wylosowany_plan_cache[idx+1] = st.session_state.wylosowany_plan_cache[idx+1], st.session_state.wylosowany_plan_cache[idx]
                        st.rerun()
                    if c_del.button("❌", key=f"del_{idx}", type="primary"):
                        st.session_state.wylosowany_plan_cache.pop(idx)
                        st.rerun()

# ZAKŁADKA 2: KREATOR MANUALNY
with tab2:
    st.subheader("Manualne dodawanie ćwiczeń")
    
    # --- NOWOŚĆ: ODDZIELNE PODZAKŁADKI I WYSZUKIWARKI (PKT 3) ---
    podzak_fizjo, podzak_gym = st.tabs(["🏥 Baza Fizjoterapia", "🏋️ Baza Siłownia"])
    
    with podzak_fizjo:
        szukaj_f = st.text_input("🔍 Wyszukaj ćwiczenie z Fizjoterapii:", key="sz_f").strip().lower()
        st.divider()
        if szukaj_f:
            for kat, lista in BAZA_FIZJO.items():
                for cw in lista:
                    if szukaj_f in cw['nazwa'].lower() or szukaj_f in cw['miesnie'].lower():
                        c1, c2 = st.columns([4, 1])
                        c1.write(f"**{cw['nazwa']}** ({kat})\n↳ [{cw['miesnie']}]")
                        if c2.button("Dodaj", key=f"df_{cw['nazwa']}"):
                            st.session_state.wylosowany_plan_cache.append((kat, cw.copy()))
                            st.toast("Dodano!")
        else:
            kat_wyb_f = st.radio("Wybierz kategorię Fizjo:", list(BAZA_FIZJO.keys()))
            for cw in BAZA_FIZJO[kat_wyb_f]:
                c1, c2 = st.columns([4, 1])
                c1.write(f"**{cw['nazwa']}** \n↳ [{cw['miesnie']}]")
                if c2.button("Dodaj", key=f"bf_{cw['nazwa']}"):
                    st.session_state.wylosowany_plan_cache.append((kat_wyb_f, cw.copy()))
                    st.toast("Dodano!")

    with podzak_gym:
        szukaj_g = st.text_input("🔍 Wyszukaj ćwiczenie z Siłowni:", key="sz_g").strip().lower()
        st.divider()
        if szukaj_g:
            for kat, lista in BAZA_SILOWNIA.items():
                for cw in lista:
                    if szukaj_g in cw['nazwa'].lower() or szukaj_g in cw['miesnie'].lower():
                        c1, c2 = st.columns([4, 1])
                        c1.write(f"**{cw['nazwa']}** ({kat})\n↳ [{cw['miesnie']}]")
                        if c2.button("Dodaj", key=f"dg_{cw['nazwa']}"):
                            st.session_state.wylosowany_plan_cache.append((f"GYM: {kat}", cw.copy()))
                            st.toast("Dodano!")
        else:
            kat_wyb_g = st.radio("Wybierz kategorię Siłownia:", list(BAZA_SILOWNIA.keys()))
            for cw in BAZA_SILOWNIA[kat_wyb_g]:
                c1, c2 = st.columns([4, 1])
                c1.write(f"**{cw['nazwa']}** \n↳ [{cw['miesnie']}]")
                if c2.button("Dodaj", key=f"bg_{cw['nazwa']}"):
                    st.session_state.wylosowany_plan_cache.append((f"GYM: {kat_wyb_g}", cw.copy()))
                    st.toast("Dodano!")

# ZAKŁADKA 3: CZAT AI GROQ
with tab3:
    st.subheader("Wirtualny Konsultant Treningowy (Llama 3)")
    if not groq_client:
        st.info("👈 Wklej klucz w panelu bocznym.")
    else:
        for msg in st.session_state.historia_wiadomosci:
            if msg["role"] != "system":
                with st.chat_message(msg["role"]): st.markdown(msg["content"])
        if prompt := st.chat_input("Pytanie do AI..."):
            with st.chat_message("user"): st.markdown(prompt)
            st.session_state.historia_wiadomosci.append({"role": "user", "content": prompt})
            with st.chat_message("assistant"):
                try:
                    completion = groq_client.chat.completions.create(model="llama-3.3-70b-versatile", messages=st.session_state.historia_wiadomosci, temperature=0.7, max_tokens=1024)
                    odp = completion.choices[0].message.content
                    st.markdown(odp)
                    st.session_state.historia_wiadomosci.append({"role": "assistant", "content": odp})
                except Exception as e: st.error(f"Błąd: {e}")

# ZAKŁADKA 4: MENEDŻER WŁASNYCH ĆWICZEŃ
with tab4:
    st.subheader("Zarządzaj własną bazą")
    with st.expander("➕ Dodaj nowe ćwiczenie", expanded=False):
        with st.form("form_dodaj_cwiczenie", clear_on_submit=True):
            kategorie = [f"FIZJO: {k}" for k in BAZA_FIZJO.keys()] + [f"GYM: {k}" for k in BAZA_SILOWNIA.keys()]
            kat_wyb = st.selectbox("Kategoria Docelowa:", kategorie)
            n_nazwa = st.text_input("Nazwa ćwiczenia:")
            n_parametry = st.text_input("Zalecenie:")
            c1, c2 = st.columns(2)
            n_czas = c1.number_input("Czas (min):", min_value=1, value=2)
            n_miesnie = c2.text_input("Anatomia:")
            n_opis = st.text_area("Opis:")
            
            if st.form_submit_button("Zapisz", type="primary"):
                if n_nazwa and n_parametry and n_miesnie and n_opis:
                    nowe_cw = {"nazwa": n_nazwa, "opis": n_opis, "czas_min": int(n_czas), "parametry": n_parametry, "miesnie": n_miesnie}
                    tb, kategoria_str = kat_wyb.split(": ")
                    dane = {"FIZJO": {}, "GYM": {}}
                    if os.path.exists(PLIK_WLASNYCH_CWICZEN):
                        try:
                            with open(PLIK_WLASNYCH_CWICZEN, "r") as f: dane = json.load(f)
                        except: pass
                    if tb not in dane: dane[tb] = {}
                    if kategoria_str not in dane[tb]: dane[tb][kategoria_str] = []
                    dane[tb][kategoria_str].append(nowe_cw)
                    with open(PLIK_WLASNYCH_CWICZEN, "w") as f: json.dump(dane, f, ensure_ascii=False, indent=4)
                    st.success("Dodano! Odśwież stronę.")
