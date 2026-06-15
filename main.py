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
        {"nazwa": "Spacer farmera z kettlebell", "opis": "INSTRUKCJA: Chwyć odważniki. Idź powolnym krokiem ze stabilnymi barkami.", "czas_min": 3, "parametry": "3 serie x 45 sekund", "miesnie": "Przedramiona, góra pleców"},
        {"nazwa": "Ślizg nerwu pośrodkowego", "opis": "INSTRUKCJA: Stojąc przy ścianie, ramię odwiedzione do 90st. Wykonuj powolny wyprost łokcia i nadgarstka, jednocześnie pochylając głowę w stronę przeciwną.", "czas_min": 2, "parametry": "15 powtórzeń na stronę", "miesnie": "Układ nerwowy obręczy barkowej"},
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
        {"nazwa": "Russian Twist", "opis": "INSTRUKCJA: Siad z uniesionymi stopami. Rotuj klatkę piersiową.", "czas_min": 2, "parametry": "3 serie x 20 skrętów", "miesnie": "Skośne brzucha"},
        {"nazwa": "Przeprosty McKenziego (Leżenie przodem)", "opis": "INSTRUKCJA: Leżenie przodem. Dłonie na wysokości barków. Powoli prostuj łokcie, unosząc samą klatkę piersiową, miednica zostaje przyklejona do maty.", "czas_min": 2, "parametry": "3x10 powtórzeń", "miesnie": "Prostownik grzbietu"},
        
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
        {"nazwa": "Krzesełko przy ścianie", "opis": "INSTRUKCJA: Oprzyj plecy o ścianę, zejdź biodrami do kąta 90 stopni. Trzymaj nieruchomo.", "czas_min": 2, "parametry": "3x30 sekund", "miesnie": "Czworogłowy uda, stabilizacja kolana"},
        {"nazwa": "Ślizg nerwu kulszowego (Neuromobilizacja)", "opis": "INSTRUKCJA: Leżenie tyłem, noga ugięta w biodrze do 90st. Prostuj kolano zadzierając palce stopy na siebie (napięcie), a opuszczając stopę obciągaj palce (luzowanie).", "czas_min": 2, "parametry": "15 powtórzeń na stronę", "miesnie": "Układ nerwowy, tylna taśma"},
        {"nazwa": "Rozciąganie zginaczy biodra w klęku", "opis": "INSTRUKCJA: Klęk jednonóż. Wypchnij miednicę mocno do przodu, zachowując wyprostowany tułów, aż poczujesz rozciąganie z przodu uda nogi zakrocznej.", "czas_min": 2, "parametry": "3x30 sekund na stronę", "miesnie": "Mięsień biodrowo-lędźwiowy, prosty uda"},
    ]
}

# ==============================================================================
# BAZA TRENINGOWA NA SIŁOWNI
# ==============================================================================
BAZA_SILOWNIA = {
    "Rozgrzewka": [
        {"nazwa": "Monster Walk z mini-bandem", "opis": "INSTRUKCJA: Załóż mini-band nad kolana, przyjmij pozycję półprzysiadu. Wykonuj szerokie kroki w bok, pilnując kolan na zewnątrz.", "czas_min": 2, "parametry": "2 min", "miesnie": "Pośladkowy średni, stabilizatory"},
        {"nazwa": "Face Pulls na wyciągu", "opis": "INSTRUKCJA: Przyciągaj linę do twarzy, rozchylając dłonie na boki i mocno spinając tył barków.", "czas_min": 2, "parametry": "3x15", "miesnie": "Naramienne tył, czworoboczny"},
        {"nazwa": "Rotacje zewnętrzne z gumą", "opis": "INSTRUKCJA: Łokieć przy boku pod kątem 90 stopni, rotuj przedramię na zewnątrz.", "czas_min": 2, "parametry": "3x12", "miesnie": "Stożek rotatorów"},
        {"nazwa": "Koci grzbiet", "opis": "INSTRUKCJA: W klęku podpartym naprzemiennie wyginaj kręgosłup w górę i w dół.", "czas_min": 2, "parametry": "10 razy", "miesnie": "Prostowniki grzbietu"},
        {"nazwa": "Skip A w miejscu", "opis": "INSTRUKCJA: Dynamiczne unoszenie kolan z zachowaniem wyprostowanej sylwetki.", "czas_min": 2, "parametry": "2 min", "miesnie": "Biodra, core"},
        {"nazwa": "Krążenia ramion z gumami", "opis": "INSTRUKCJA: Obszerne krążenia z lekką taśmą w dłoniach.", "czas_min": 2, "parametry": "20 razy", "miesnie": "Barki"},
        {"nazwa": "Pajacyki", "opis": "INSTRUKCJA: Klasyczne skoki z odwodzeniem ramion i nóg.", "czas_min": 2, "parametry": "2 min", "miesnie": "Całe ciało"},
        {"nazwa": "Wymachy nóg w przód", "opis": "INSTRUKCJA: Stojąc, dynamicznie wymachuj prostą nogą w przód.", "czas_min": 2, "parametry": "10 na stronę", "miesnie": "Biodra, dwugłowe"},
        {"nazwa": "Przysiady z masą ciała", "opis": "INSTRUKCJA: Płynne przysiady bez obciążenia, pilnując ustawienia stóp.", "czas_min": 2, "parametry": "20 razy", "miesnie": "Nogi"},
        {"nazwa": "Pompki przy ścianie", "opis": "INSTRUKCJA: Wstępna aktywacja klatki przez odpychanie się od ściany.", "czas_min": 2, "parametry": "15 razy", "miesnie": "Klatka, triceps"},
        {"nazwa": "Wykroki z rotacją", "opis": "INSTRUKCJA: Wykonaj wykrok i skręt tułowia w stronę nogi wykrocznej.", "czas_min": 3, "parametry": "5 na stronę", "miesnie": "Biodra, kręgosłup"},
        {"nazwa": "Skakanka", "opis": "INSTRUKCJA: Skoki przez skakankę w stałym tempie.", "czas_min": 3, "parametry": "3 min", "miesnie": "Łydki, serce"},
        {"nazwa": "Plank", "opis": "INSTRUKCJA: Podpór przodem na przedramionach, ciało w linii prostej.", "czas_min": 2, "parametry": "1 min", "miesnie": "Core"},
        {"nazwa": "Bird-Dog", "opis": "INSTRUKCJA: W klęku podpartym unoś przeciwną rękę i nogę.", "czas_min": 2, "parametry": "10 na stronę", "miesnie": "Core"},
        {"nazwa": "Wiosłowanie taśmą", "opis": "INSTRUKCJA: Zaczep taśmę o słup i przyciągaj końce do bioder.", "czas_min": 2, "parametry": "20 razy", "miesnie": "Plecy"},
        {"nazwa": "Mostki biodrowe", "opis": "INSTRUKCJA: Leżenie tyłem, unoś biodra poprzez napięcie pośladków.", "czas_min": 2, "parametry": "20 razy", "miesnie": "Pośladki"},
        {"nazwa": "Spacer w podporze", "opis": "INSTRUKCJA: Z pozycji stojącej przejdź dłońmi do podporu i wróć.", "czas_min": 2, "parametry": "10 razy", "miesnie": "Core, ramiona"},
        {"nazwa": "Krążenia bioder", "opis": "INSTRUKCJA: Obszerne ruchy miednicą w staniu.", "czas_min": 1, "parametry": "1 min", "miesnie": "Biodra"},
        {"nazwa": "Rowerek - lekki opór", "opis": "INSTRUKCJA: Spokojne pedałowanie w celu rozgrzania stawów.", "czas_min": 3, "parametry": "3 min", "miesnie": "Nogi"},
        {"nazwa": "Wspinaczka górska (Mountain Climbers)", "opis": "INSTRUKCJA: W podporze dynamicznie przyciągaj kolana do klatki.", "czas_min": 2, "parametry": "1 min", "miesnie": "Core, nogi"}
    ],
    "Klatka piersiowa": [
        {"nazwa": "Wyciskanie sztangi płasko", "opis": "POZYCJA: Połóż się na ławce, stopy stabilnie na podłożu, łopatki ściągnięte. RUCH: Opuść sztangę do dolnej części klatki, a następnie wyciśnij dynamicznie w górę. UWAGA: Nie odrywaj pośladków od ławki, kontroluj ciężar przy opuszczaniu.", "czas_min": 3, "parametry": "4x8", "miesnie": "Klatka, triceps"},
        {"nazwa": "Wyciskanie hantli skos dodatni", "opis": "POZYCJA: Ławka ustawiona pod kątem 30 stopni. Stopy stabilne. RUCH: Wyciskaj hantle nad klatkę, zbliżając je do siebie w górnej fazie. UWAGA: Nie przeprostowuj łokci w szczycie ruchu.", "czas_min": 3, "parametry": "4x10", "miesnie": "Góra klatki, barki"},
        {"nazwa": "Dipy (pompki na poręczach)", "opis": "POZYCJA: Podeprzyj się na poręczach, tułów pochylony w przód. RUCH: Uginaj ramiona, aż poczujesz rozciągnięcie klatki, następnie wróć do pełnego wyprostu. UWAGA: Prowadź łokcie szerzej niż przy tricepsie.", "czas_min": 3, "parametry": "3x10", "miesnie": "Klatka, triceps"},
        {"nazwa": "Rozpiętki z hantlami", "opis": "POZYCJA: Leżąc płasko, hantle nad klatką, lekko ugięte łokcie. RUCH: Otwieraj ramiona szeroko na boki, aż poczujesz mocne rozciągnięcie, wróć do pozycji początkowej. UWAGA: Ruch wykonuj w stawie barkowym, nie zmieniaj kąta ugięcia łokci.", "czas_min": 2, "parametry": "3x12", "miesnie": "Klatka"},
        {"nazwa": "Wyciskanie hantli płasko", "opis": "POZYCJA: Leżenie na ławce płaskiej. RUCH: Wyciskaj hantle w pełnym zakresie ruchu, dbając o symetrię. UWAGA: Kontroluj tor ruchu hantli, aby nie traciły stabilności.", "czas_min": 3, "parametry": "4x10", "miesnie": "Klatka"},
        {"nazwa": "Pompki klasyczne", "opis": "POZYCJA: Podpora przodem, dłonie szerzej niż barki. RUCH: Obniż tułów do ziemi, zachowując linię prostą, wróć do góry. UWAGA: Napinaj brzuch, aby uniknąć zapadania się lędźwi.", "czas_min": 2, "parametry": "3xMax", "miesnie": "Klatka, triceps, core"},
        {"nazwa": "Brama (Cable Crossover)", "opis": "POZYCJA: Stań w wykroku w środku bramy. RUCH: Ściągaj uchwyty wyciągu w dół i do środka, spinając klatkę. UWAGA: Prowadź ruch po łuku, nie wyginaj kręgosłupa.", "czas_min": 3, "parametry": "3x15", "miesnie": "Klatka"},
        {"nazwa": "Wyciskanie Hammer Strength", "opis": "POZYCJA: Siedząc na maszynie, chwyć uchwyty. RUCH: Wypychaj uchwyty przed siebie. UWAGA: Nie odrywaj pleców od oparcia podczas wysiłku.", "czas_min": 3, "parametry": "3x12", "miesnie": "Klatka, triceps"},
        {"nazwa": "Pompki na podwyższeniu nóg", "opis": "POZYCJA: Stopy na ławce, dłonie na ziemi. RUCH: Pompka klasyczna. UWAGA: Skupienie na górnej części klatki, pilnuj bioder.", "czas_min": 2, "parametry": "3x12", "miesnie": "Góra klatki, barki"},
        {"nazwa": "Wyciskanie w maszynie Smitha", "opis": "POZYCJA: Leżenie na ławce pod gryfem maszyny. RUCH: Wyciskaj gryf, który porusza się po szynie. UWAGA: Dopasuj pozycję ławki, aby gryf trafiał na linię mostka.", "czas_min": 3, "parametry": "4x10", "miesnie": "Klatka"},
        {"nazwa": "Floor Press (wyciskanie z podłogi)", "opis": "POZYCJA: Leżenie na plecach, sztanga nad klatką. RUCH: Opuść sztangę, aż tricepsy dotkną podłogi, następnie wyciśnij. UWAGA: Zatrzymuj ruch na podłodze, nie odbijaj.", "czas_min": 3, "parametry": "3x10", "miesnie": "Klatka, triceps"},
        {"nazwa": "Pompki diamentowe", "opis": "POZYCJA: Podpór przodem, dłonie blisko siebie (kciuki i palce wskazujące stykają się). RUCH: Pompka. UWAGA: Skupienie na środku klatki i pracy tricepsów.", "czas_min": 2, "parametry": "3x10", "miesnie": "Środek klatki, triceps"},
        {"nazwa": "Landmine Press (jednorącz)", "opis": "POZYCJA: Stojąc, koniec sztangi zamocowany w rogu. RUCH: Wyciskaj koniec sztangi jedną ręką przed siebie. UWAGA: Stabilizuj tułów napięciem brzucha.", "czas_min": 3, "parametry": "3x12", "miesnie": "Góra klatki, barki"},
        {"nazwa": "Wyciskanie hantla jednorącz leżąc", "opis": "POZYCJA: Leżąc, jedna ręka pracuje, druga stabilizuje. RUCH: Wyciskaj hantel w górę. UWAGA: Utrzymuj stabilną pozycję barków na ławce.", "czas_min": 3, "parametry": "3x12", "miesnie": "Klatka"},
        {"nazwa": "Maszyna Butterfly (motylki)", "opis": "POZYCJA: Siedząc, przedramiona na poduszkach maszyny. RUCH: Zbliżaj ramiona do siebie. UWAGA: Nie przesuwaj ramion za wysoko, łokcie na linii barków.", "czas_min": 2, "parametry": "3x15", "miesnie": "Klatka"},
        {"nazwa": "Wyciskanie sztangi skos ujemny", "opis": "POZYCJA: Ławka ustawiona pod skosem w dół. RUCH: Wyciskaj sztangę. UWAGA: Kontroluj tempo, aby sztanga nie spadła na szyję.", "czas_min": 3, "parametry": "3x10", "miesnie": "Dół klatki"},
        {"nazwa": "Pompki z taśmą oporową", "opis": "POZYCJA: Taśma przełożona przez plecy, dłonie trzymają końce. RUCH: Pompka z dodatkowym oporem. UWAGA: Pilnuj, by taśma nie ześlizgnęła się z pleców.", "czas_min": 2, "parametry": "3x10", "miesnie": "Klatka"},
        {"nazwa": "Pullover z hantlem", "opis": "POZYCJA: Leżąc prostopadle do ławki. RUCH: Przenoś hantel zza głowy nad klatkę po łuku. UWAGA: Rozciągaj klatkę w bezpiecznym zakresie.", "czas_min": 2, "parametry": "3x12", "miesnie": "Klatka, najszerszy"},
        {"nazwa": "Wyciskanie hantli chwytem neutralnym", "opis": "POZYCJA: Leżenie płasko. RUCH: Wyciskaj hantle, trzymając dłonie skierowane do siebie. UWAGA: Bardzo stabilna praca dla barków i klatki.", "czas_min": 3, "parametry": "3x10", "miesnie": "Klatka"},
        {"nazwa": "Wyciskanie sztangi wąsko", "opis": "POZYCJA: Chwyt sztangi na szerokość barków. RUCH: Wyciskaj, prowadząc łokcie blisko tułowia. UWAGA: Mocna praca tricepsa, ale też środek klatki.", "czas_min": 3, "parametry": "3x8", "miesnie": "Środek klatki, triceps"}
    ],
    "Plecy": [
        {"nazwa": "Podciąganie na drążku nachwytem", "opis": "POZYCJA: Chwyć drążek szerzej niż barki, ramiona wyprostowane. RUCH: Podciągnij klatkę piersiową do drążka, prowadząc łokcie w dół i do tyłu. UWAGA: Unikaj huśtania ciałem; ruch powinien być kontrolowany w obie strony.", "czas_min": 3, "parametry": "3 serie x Max", "miesnie": "Najszerszy grzbietu, obły większy"},
        {"nazwa": "Wiosłowanie sztangą w opadzie", "opis": "POZYCJA: Stopy na szerokość bioder, kolana lekko ugięte. Pochyl tułów do kąta 45 stopni, plecy proste. RUCH: Przyciągaj sztangę do dolnej części brzucha, ściągając łopatki. UWAGA: Kręgosłup musi być w linii prostej przez cały czas.", "czas_min": 3, "parametry": "4 serie x 10 powtórzeń", "miesnie": "Plecy, prostownik grzbietu, biceps"},
        {"nazwa": "Martwy ciąg klasyczny", "opis": "POZYCJA: Stopy na szerokość bioder, sztanga blisko piszczeli. RUCH: Podnieś sztangę, prostując biodra i kolana jednocześnie, wypychając klatkę w przód. UWAGA: Nie wyginaj kręgosłupa w łuk; ciężar prowadź jak najbliżej nóg.", "czas_min": 4, "parametry": "4 serie x 6 powtórzeń", "miesnie": "Prostowniki grzbietu, pośladki, plecy"},
        {"nazwa": "Wiosłowanie hantlem jednorącz", "opis": "POZYCJA: Jedna ręka i kolano na ławce, plecy równolegle do podłoża. RUCH: Przyciągaj hantel w stronę biodra, łokieć prowadź blisko tułowia. UWAGA: Unikaj rotacji tułowia podczas ruchu.", "czas_min": 3, "parametry": "3 serie x 12 na stronę", "miesnie": "Najszerszy grzbietu, czworoboczny"},
        {"nazwa": "Ściąganie drążka wyciągu górnego", "opis": "POZYCJA: Usiądź, zablokuj uda pod wałkiem. RUCH: Przyciągaj drążek do górnej części klatki piersiowej, lekko odchylając się w tył. UWAGA: Nie bujaj ciałem, pracuj świadomie mięśniami pleców.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Najszerszy grzbietu"},
        {"nazwa": "Wiosłowanie na maszynie (Hammer)", "opis": "POZYCJA: Siedząc na maszynie, chwyć uchwyty. RUCH: Przyciągaj uchwyty do brzucha, mocno spinając łopatki w szczycie. UWAGA: Nie odrywaj klatki od oparcia maszyny.", "czas_min": 3, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Mięśnie pleców"},
        {"nazwa": "Face Pulls z linką", "opis": "POZYCJA: Stań przed wyciągiem górnym. RUCH: Przyciągaj linkę do czoła, rozchylając dłonie na boki i rotując ramiona na zewnątrz. UWAGA: Łokcie muszą znajdować się wyżej niż dłonie.", "czas_min": 2, "parametry": "3 serie x 15 powtórzeń", "miesnie": "Czworoboczny, tył barków"},
        {"nazwa": "Wiosłowanie na pasach TRX", "opis": "POZYCJA: Chwyć uchwyty, odchyl ciało w tył w pozycji deski. RUCH: Przyciągaj klatkę do dłoni, pracując łopatkami. UWAGA: Napnij pośladki i brzuch, aby ciało nie zwisało.", "czas_min": 2, "parametry": "3 serie x 15 powtórzeń", "miesnie": "Plecy, core"},
        {"nazwa": "Pullover hantlem (leżąc)", "opis": "POZYCJA: Leżenie w poprzek ławki, głowa poza nią. RUCH: Przenoś hantel zza głowy nad klatkę po łuku. UWAGA: Ruch wykonuj powoli, pilnując stabilności obręczy barkowej.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Najszerszy grzbietu, klatka"},
        {"nazwa": "Wiosłowanie kettlem", "opis": "POZYCJA: Szeroki rozkrok, kettlebell na ziemi między stopami. RUCH: Podciągaj kettlebell do brzucha w opadzie. UWAGA: Plecy muszą pozostać proste, nie zaokrąglaj lędźwi.", "czas_min": 3, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Plecy"},
        {"nazwa": "Ściąganie uchwytu V na wyciągu", "opis": "POZYCJA: Siedząc, chwyt wąski. RUCH: Przyciągaj uchwyt V do splotu słonecznego. UWAGA: Prowadź ruch płynnie, bez szarpania ciężarem.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Najszerszy grzbietu"},
        {"nazwa": "Shrugsy (wzruszanie barkami)", "opis": "POZYCJA: Stojąc, hantle w dłoniach wzdłuż tułowia. RUCH: Wzruszaj barkami w górę, jakbyś chciał dotknąć nimi uszu. UWAGA: Nie rotuj barkami, ruch tylko góra-dół.", "czas_min": 2, "parametry": "4 serie x 12 powtórzeń", "miesnie": "Czworoboczny"},
        {"nazwa": "Wiosłowanie taśmą oporową", "opis": "POZYCJA: Taśma zaczepiona o stabilny słup. RUCH: Przyciągaj końce taśmy do bioder, pilnując postawy. UWAGA: Kontroluj fazę powrotu (odpuszczania) taśmy.", "czas_min": 2, "parametry": "3 serie x 20 powtórzeń", "miesnie": "Plecy"},
        {"nazwa": "Podciąganie podchwytem", "opis": "POZYCJA: Chwyt drążka od dołu. RUCH: Podciągaj klatkę w górę, mocno angażując mięśnie ramion i pleców. UWAGA: Pełen zakres ruchu w dół.", "czas_min": 3, "parametry": "3 serie x Max", "miesnie": "Plecy, biceps"},
        {"nazwa": "Dzień dobry (Good Mornings)", "opis": "POZYCJA: Sztanga na karku, stopy na szerokość bioder. RUCH: Skłon tułowia z prostymi (lekko ugiętymi) nogami, biodra w tył. UWAGA: Ruch tylko do momentu, gdy poczujesz rozciąganie dwugłowych.", "czas_min": 3, "parametry": "3 serie x 10 powtórzeń", "miesnie": "Prostowniki grzbietu, dwugłowe"},
        {"nazwa": "Wiosłowanie na ławce skosnej", "opis": "POZYCJA: Leżenie klatką na ławce skośnej. RUCH: Wiosłuj hantlami do boków ciała. UWAGA: Nie odrywaj klatki od ławeczki.", "czas_min": 3, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Plecy, tył barków"},
        {"nazwa": "High Pulls (sztanga)", "opis": "POZYCJA: Sztanga w dłoniach, stopy na szerokość bioder. RUCH: Dynamiczne wyprostowanie bioder z jednoczesnym podciągnięciem sztangi do brody. UWAGA: Użyj siły nóg do zainicjowania ruchu.", "czas_min": 3, "parametry": "3 serie x 8 powtórzeń", "miesnie": "Plecy, barki"},
        {"nazwa": "Wiosłowanie na wyciągu dolnym", "opis": "POZYCJA: Siedząc z wyprostowanymi nogami. RUCH: Przyciągaj rączkę do brzucha, wypinając klatkę w przód. UWAGA: Unikaj odchylania się nadmiernie w tył.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Plecy"},
        {"nazwa": "T-Bar Row (wiosłowanie T-Bar)", "opis": "POZYCJA: W opadzie, chwyt za uchwyty maszyny. RUCH: Przyciągaj ciężar do klatki. UWAGA: Pilnuj, by plecy były płaskie, nie zaokrąglaj się.", "czas_min": 3, "parametry": "4 serie x 10 powtórzeń", "miesnie": "Plecy"},
        {"nazwa": "Spacer farmera", "opis": "POZYCJA: Wyprostowana sylwetka, hantle w dłoniach. RUCH: Marsz krótkimi krokami, stabilizując tułów. UWAGA: Nie pozwól, aby ciężar ciągnął ramiona w dół, trzymaj barki aktywnie.", "czas_min": 3, "parametry": "3 serie x 30m", "miesnie": "Plecy, chwyt, core"}
    ],
    "Ręce": [
        {"nazwa": "Uginanie sztangi łamanej", "opis": "POZYCJA: Stojąc, chwyć sztangę podchwytem na szerokość barków. RUCH: Uginaj ramiona w łokciach, prowadząc gryf w stronę klatki piersiowej. UWAGA: Łokcie muszą pozostać w jednej pozycji przy tułowiu, nie wykonuj zamachu tułowiem.", "czas_min": 2, "parametry": "3x12", "miesnie": "Biceps, m. ramienny"},
        {"nazwa": "Uginanie hantli z supinacją", "opis": "POZYCJA: Stojąc z hantlami w dłoniach, chwyt neutralny. RUCH: Uginaj ramię, jednocześnie obracając nadgarstek na zewnątrz (dłoń do góry). UWAGA: Rotacja następuje w trakcie ruchu, nie na końcu.", "czas_min": 2, "parametry": "3x12", "miesnie": "Biceps"},
        {"nazwa": "Wyciskanie francuskie sztangi", "opis": "POZYCJA: Leżąc na ławce, sztanga nad czołem. RUCH: Zginaj łokcie, opuszczając sztangę do czoła, następnie prostuj ramiona. UWAGA: Łokcie nie mogą rozchodzić się na boki.", "czas_min": 2, "parametry": "3x10", "miesnie": "Triceps"},
        {"nazwa": "Prostowanie linek na wyciągu", "opis": "POZYCJA: Stojąc, łokcie przyklejone do talii. RUCH: Prostuj ramiona, rozciągając linkę na dole. UWAGA: Skup się na pełnym wyproście w łokciu bez ruchu barków.", "czas_min": 2, "parametry": "3x15", "miesnie": "Triceps"},
        {"nazwa": "Uginanie hantli młotkowe", "opis": "POZYCJA: Stojąc, hantle chwytem neutralnym (dłonie do siebie). RUCH: Uginaj ramię bez rotacji nadgarstka. UWAGA: Trzymaj nadgarstek sztywno przez cały czas.", "czas_min": 2, "parametry": "3x12", "miesnie": "Mięsień ramienny, biceps"},
        {"nazwa": "Dipy na ławeczce", "opis": "POZYCJA: Ręce na ławeczce za plecami, stopy na podłożu. RUCH: Obniżaj biodra, uginając łokcie do kąta prostego, potem wróć. UWAGA: Plecy muszą prowadzić blisko ławki.", "czas_min": 2, "parametry": "3x15", "miesnie": "Triceps"},
        {"nazwa": "Uginanie na modlitewniku", "opis": "POZYCJA: Przedramiona oparte o oparcie modlitewnika. RUCH: Uginaj ramię, pilnując, by łokieć nie odrywał się od oparcia. UWAGA: Nie wykonuj ruchu w pełnym wyproście, aby chronić stawy.", "czas_min": 2, "parametry": "3x12", "miesnie": "Biceps"},
        {"nazwa": "Prostowanie hantla za głowę", "opis": "POZYCJA: Siedząc lub stojąc, hantel oburącz nad głową. RUCH: Zginaj łokcie, opuszczając hantel za kark, potem prostuj. UWAGA: Łokcie powinny być skierowane do przodu/góry, nie na boki.", "czas_min": 2, "parametry": "3x12", "miesnie": "Triceps (głowa długa)"},
        {"nazwa": "Uginanie z gumą oporową", "opis": "POZYCJA: Stań na środku gumy, końce w dłoniach. RUCH: Uginaj ramiona z gumą. UWAGA: Kontroluj fazę powrotu (odpuszczania) gumy, nie pozwól, by szarpnęła ręką.", "czas_min": 2, "parametry": "3x20", "miesnie": "Biceps"},
        {"nazwa": "Zottman Curl", "opis": "POZYCJA: Stojąc z hantlami. RUCH: Ugnij ręce z supinacją (dłoń do góry), w szczycie obróć nadgarstki i opuszczaj hantle chwytem pronowanym (dłonie w dół). UWAGA: Płynna rotacja w szczycie.", "czas_min": 2, "parametry": "3x10", "miesnie": "Biceps, przedramiona"},
        {"nazwa": "Triceps podchwytem (wyciąg)", "opis": "POZYCJA: Stojąc przy wyciągu, chwyt jednorącz podchwytem. RUCH: Prostuj ramię w dół. UWAGA: Łokieć jest osią obrotu, trzymaj go nieruchomo przy ciele.", "czas_min": 2, "parametry": "3x12", "miesnie": "Triceps"},
        {"nazwa": "Spider Curl (na ławce skośnej)", "opis": "POZYCJA: Leżąc klatką na ławce skośnej, ramiona pionowo w dół. RUCH: Uginaj hantle w górę. UWAGA: Nie wykonuj zamachu łokciami, trzymaj je w jednej linii.", "czas_min": 2, "parametry": "3x12", "miesnie": "Biceps"},
        {"nazwa": "Pompki diamentowe", "opis": "POZYCJA: Dłonie ułożone w kształt diamentu. RUCH: Pompka. UWAGA: Podczas ruchu łokcie prowadź blisko tułowia.", "czas_min": 2, "parametry": "3x12", "miesnie": "Triceps"},
        {"nazwa": "Uginanie koncentryczne", "opis": "POZYCJA: W siadzie, łokieć oparty o wewnętrzną stronę kolana. RUCH: Uginaj ramię. UWAGA: Nie wspomagaj się drugą ręką.", "czas_min": 2, "parametry": "3x12", "miesnie": "Biceps"},
        {"nazwa": "Skull Crushers z hantlami", "opis": "POZYCJA: Leżenie, hantle nad głową. RUCH: Opuszczaj hantle po bokach głowy, prostuj. UWAGA: Stabilne nadgarstki.", "czas_min": 2, "parametry": "3x10", "miesnie": "Triceps"},
        {"nazwa": "Uginanie z kettlem", "opis": "POZYCJA: Chwyt za kulę kettla oburącz. RUCH: Uginaj ramiona do klatki. UWAGA: Stabilna pozycja kręgosłupa.", "czas_min": 2, "parametry": "3x12", "miesnie": "Biceps"},
        {"nazwa": "Prostowanie za głowę z gumą", "opis": "POZYCJA: Guma zaczepiona nisko, za plecami. RUCH: Prostuj ręce nad głowę. UWAGA: Utrzymuj stabilny korpus.", "czas_min": 2, "parametry": "3x15", "miesnie": "Triceps"},
        {"nazwa": "Uginanie nadgarstków sztangą", "opis": "POZYCJA: Przedramiona oparte o ławkę, dłonie poza nią. RUCH: Uginaj nadgarstki. UWAGA: Małe, kontrolowane ruchy.", "czas_min": 2, "parametry": "3x15", "miesnie": "Przedramiona"},
        {"nazwa": "System 21 (Biceps)", "opis": "POZYCJA: Stojąc. RUCH: 7 powtórzeń do połowy w górę, 7 od połowy do góry, 7 pełnych. UWAGA: Bardzo męczące, pilnuj techniki przy każdym odcinku.", "czas_min": 3, "parametry": "3x21", "miesnie": "Biceps"},
        {"nazwa": "Wyprosty ramion jednorącz", "opis": "POZYCJA: Stojąc przy wyciągu. RUCH: Prostuj ramię do pełnego wyprostu. UWAGA: Pełne wyczucie pracy tricepsa.", "czas_min": 2, "parametry": "3x15", "miesnie": "Triceps"}
    ],
    "Nogi": [
        {"nazwa": "Przysiad ze sztangą na plecach", "opis": "POZYCJA: Stopy na szerokość barków, sztanga na mięśniach czworobocznych (nie na kręgach). RUCH: Wypchnij biodra w tył i zginaj kolana, utrzymując proste plecy. UWAGA: Kolana powinny podążać za linią palców, nie schodź do środka.", "czas_min": 4, "parametry": "4x8", "miesnie": "Czworogłowe, pośladki, przywodziciele"},
        {"nazwa": "Martwy ciąg na prostych nogach", "opis": "POZYCJA: Sztanga w dłoniach, stopy na szerokość bioder, kolana lekko ugięte. RUCH: Wypychaj biodra mocno w tył, prowadząc sztangę blisko nóg do połowy piszczeli. UWAGA: Kręgosłup musi być idealnie prosty, nie garb się.", "czas_min": 3, "parametry": "4x10", "miesnie": "Dwugłowe uda, pośladki, prostowniki"},
        {"nazwa": "Wykroki z hantlami", "opis": "POZYCJA: Stojąc, hantle w dłoniach wzdłuż tułowia. RUCH: Wykonaj duży krok w przód, uginając oba kolana do kąta 90 stopni. UWAGA: Tułów trzymaj pionowo, nie pochylaj się do przodu.", "czas_min": 3, "parametry": "3x12 na nogę", "miesnie": "Czworogłowe, pośladki"},
        {"nazwa": "Wypychanie na suwnicy", "opis": "POZYCJA: Siedząc na maszynie, stopy na platformie. RUCH: Wypychaj platformę nogami, unikając przeprostu w kolanach. UWAGA: Nie odrywaj lędźwi od oparcia maszyny.", "czas_min": 3, "parametry": "4x12", "miesnie": "Czworogłowe, pośladki"},
        {"nazwa": "Wyprosty nóg na maszynie", "opis": "POZYCJA: Siedząc, wałek pod stawami skokowymi. RUCH: Prostuj kolana, napinając mięśnie ud. UWAGA: Ruch wykonuj płynnie, bez gwałtownych szarpnięć ciężarem.", "czas_min": 2, "parametry": "3x15", "miesnie": "Czworogłowe"},
        {"nazwa": "Uginanie nóg na maszynie", "opis": "POZYCJA: Leżąc przodem, wałek nad piętami. RUCH: Zginaj kolana, przyciągając pięty do pośladków. UWAGA: Biodra powinny cały czas przylegać do ławki maszyny.", "czas_min": 2, "parametry": "3x15", "miesnie": "Dwugłowe uda"},
        {"nazwa": "Przysiad Goblet (z kettlem)", "opis": "POZYCJA: Kettlebell trzymany oburącz przy klatce piersiowej. RUCH: Wykonuj głęboki przysiad. UWAGA: Kettlebell pomaga utrzymać pionową sylwetkę, nie garb się.", "czas_min": 3, "parametry": "3x12", "miesnie": "Czworogłowe, core"},
        {"nazwa": "Przysiad bułgarski", "opis": "POZYCJA: Jedna noga oparta o ławkę w tyle. RUCH: Przysiad na nodze wykrocznej. UWAGA: Bardzo wymagające ćwiczenie na stabilizację, pilnuj kolana.", "czas_min": 3, "parametry": "3x10 na nogę", "miesnie": "Czworogłowe, pośladki"},
        {"nazwa": "Wspięcia na palce (maszyna)", "opis": "POZYCJA: Stopy na krawędzi platformy. RUCH: Unoszenie pięt jak najwyżej. UWAGA: Pełen zakres ruchu – rozciągnięcie w dole i spięcie w górze.", "czas_min": 2, "parametry": "4x20", "miesnie": "Łydki"},
        {"nazwa": "Kettlebell Swing", "opis": "POZYCJA: Stojąc, kettlebell między nogami. RUCH: Dynamiczny wymach bioder w przód. UWAGA: Siła ruchu płynie z bioder, a nie z barków.", "czas_min": 3, "parametry": "3x15", "miesnie": "Dwugłowe, pośladki, core"},
        {"nazwa": "Wykroki z gumą oporową", "opis": "POZYCJA: Guma pod przednią stopą, końce w dłoniach. RUCH: Klasyczny wykrok z dodatkowym oporem gumy. UWAGA: Kontroluj tempo fazy ekscentrycznej (opuszczania).", "czas_min": 3, "parametry": "3x12", "miesnie": "Czworogłowe, pośladki"},
        {"nazwa": "Przysiad Sumo", "opis": "POZYCJA: Bardzo szeroki rozkrok, palce stóp na zewnątrz. RUCH: Przysiad z ciężarem w dłoniach. UWAGA: Plecy proste, większy nacisk na wewnętrzną część ud.", "czas_min": 3, "parametry": "3x12", "miesnie": "Przywodziciele, czworogłowe"},
        {"nazwa": "Zakroki", "opis": "POZYCJA: Stojąc, wykonaj krok w tył. RUCH: Obniżaj biodra, aż kolano tylnej nogi dotknie ziemi. UWAGA: Zachowaj stabilny tułów, nie chwiej się.", "czas_min": 3, "parametry": "3x10 na nogę", "miesnie": "Czworogłowe, pośladki"},
        {"nazwa": "Wyprosty nóg w siadzie z gumą", "opis": "POZYCJA: Siedząc, guma zaczepiona o nogę krzesła i staw skokowy. RUCH: Prostuj nogę. UWAGA: Powolna faza powrotu, maksymalne napięcie uda.", "czas_min": 2, "parametry": "3x20", "miesnie": "Czworogłowe"},
        {"nazwa": "Martwy ciąg na jednej nodze", "opis": "POZYCJA: Stojąc na jednej nodze, hantel w dłoni. RUCH: Skłon z jednoczesnym wyprostowaniem drugiej nogi w tył. UWAGA: Ćwiczenie na równowagę, skup się na punkcie przed sobą.", "czas_min": 3, "parametry": "3x10 na nogę", "miesnie": "Dwugłowe, pośladki, core"},
        {"nazwa": "Wchodzenie na skrzynię (Step-up)", "opis": "POZYCJA: Stojąc przed skrzynią. RUCH: Wejdź na skrzynię, prostując nogę, potem wróć. UWAGA: Cała stopa musi znajdować się na skrzyni.", "czas_min": 3, "parametry": "3x12 na nogę", "miesnie": "Pośladki, czworogłowe"},
        {"nazwa": "Sissy Squat", "opis": "POZYCJA: Stojąc, trzymając się stabilnego punktu. RUCH: Odchyl tułów w tył, wypychając kolana daleko w przód. UWAGA: Tylko dla osób bez problemów ze stawami kolanowymi.", "czas_min": 2, "parametry": "3x10", "miesnie": "Czworogłowe"},
        {"nazwa": "Wykroki chodzone z hantlami", "opis": "POZYCJA: Stojąc z hantlami. RUCH: Wykonuj naprzemienne wykroki idąc przed siebie. UWAGA: Pilnuj stabilności, nie pozwól kolanom uciekać do wewnątrz.", "czas_min": 3, "parametry": "3x20m", "miesnie": "Czworogłowe, pośladki"},
        {"nazwa": "Przysiad z gumą nad kolanami", "opis": "POZYCJA: Guma nad kolanami. RUCH: Wykonuj przysiad, cały czas rozpychając gumę kolanami. UWAGA: To ćwiczenie uczy prawidłowej pracy kolan.", "czas_min": 3, "parametry": "3x15", "miesnie": "Nogi, pośladki średnie"},
        {"nazwa": "Wspięcia na palce bez obciążenia", "opis": "POZYCJA: Stojąc, ręce na biodrach. RUCH: Dynamiczne wspięcia na palce. UWAGA: Duża objętość dla poprawy wytrzymałości mięśni łydki.", "czas_min": 2, "parametry": "3x30", "miesnie": "Łydki"}
    ],
   "Pośladki": [
        {"nazwa": "Hip Thrust (ze sztangą)", "opis": "POZYCJA: Łopatki oparte o stabilną ławkę, sztanga na biodrach. RUCH: Wypchnij biodra w górę, spinając pośladki w szczycie ruchu. UWAGA: Nie wyginaj odcinka lędźwiowego – ruch zachodzi w biodrach, nie w plecach.", "czas_min": 4, "parametry": "4x10", "miesnie": "Pośladkowy wielki"},
        {"nazwa": "Odwodzenie nogi na wyciągu", "opis": "POZYCJA: Stojąc, opaska na kostce, trzymaj się wyciągu. RUCH: Odprowadzaj prostą nogę w tył, kontrolując napięcie pośladka. UWAGA: Tułów stabilny, nie wykonuj zamachu tułowiem.", "czas_min": 2, "parametry": "3x15", "miesnie": "Pośladkowy wielki"},
        {"nazwa": "Glute Bridge z gumą", "opis": "POZYCJA: Leżenie tyłem, guma nad kolanami. RUCH: Unoś biodra, rozpychając kolana na zewnątrz przeciwko oporowi gumy. UWAGA: Stopy mocno wciśnięte w ziemię.", "czas_min": 2, "parametry": "3x20", "miesnie": "Pośladkowy średni, wielki"},
        {"nazwa": "Clamshell (Muszelka)", "opis": "POZYCJA: Leżenie na boku, kolana ugięte 90st. RUCH: Unoś górne kolano, nie rozłączając stóp. UWAGA: Miednica musi być nieruchoma, nie odchylaj się w tył.", "czas_min": 2, "parametry": "3x20", "miesnie": "Pośladkowy średni"},
        {"nazwa": "Martwy ciąg Rumuński", "opis": "POZYCJA: Sztanga/hantle w dłoniach, lekko ugięte kolana. RUCH: Wypychaj biodra mocno w tył, aż poczujesz rozciąganie pośladków. UWAGA: Plecy proste, ciężar blisko nóg.", "czas_min": 3, "parametry": "3x12", "miesnie": "Pośladki, dwugłowe"},
        {"nazwa": "Kabel kickbacks (wypychanie)", "opis": "POZYCJA: W klęku podpartym, opaska wyciągu na stopie. RUCH: Wypychaj piętę w górę w stronę sufitu. UWAGA: Nie wyginaj nadmiernie kręgosłupa w lędźwiach.", "czas_min": 3, "parametry": "3x15", "miesnie": "Pośladkowy wielki"},
        {"nazwa": "Wspięcia na skrzynię (Step-up)", "opis": "POZYCJA: Stań przed skrzynią. RUCH: Wejdź na skrzynię, mocno napinając pośladek nogi, która wykonuje pracę. UWAGA: Nie odpychaj się nogą znajdującą się na ziemi.", "czas_min": 3, "parametry": "3x10 na nogę", "miesnie": "Pośladki, nogi"},
        {"nazwa": "Fire Hydrant z gumą", "opis": "POZYCJA: W klęku podpartym, guma nad kolanami. RUCH: Odwodź nogę ugiętą w kolanie do boku (jak pies przy hydrancie). UWAGA: Ruch w biodrze, tułów nieruchomo.", "czas_min": 2, "parametry": "3x15", "miesnie": "Pośladkowy średni"},
        {"nazwa": "Przysiad Sumo (z kettlem)", "opis": "POZYCJA: Bardzo szeroki rozkrok, palce stóp na zewnątrz. RUCH: Przysiad z kettlem między nogami. UWAGA: Pilnuj, aby kolana nie schodziły do środka.", "czas_min": 3, "parametry": "3x12", "miesnie": "Pośladki, przywodziciele"},
        {"nazwa": "Abdukcja na maszynie", "opis": "POZYCJA: Siedząc na maszynie, rozpychaj poduszki na zewnątrz. RUCH: Powolny ruch rozwodzenia nóg. UWAGA: W fazie powrotu kontroluj ciężar, nie pozwól mu uderzyć o stos.", "czas_min": 2, "parametry": "3x15", "miesnie": "Pośladkowy średni"},
        {"nazwa": "Hip Thrust na 1 nodze", "opis": "POZYCJA: Jedna noga w górze, druga na podłożu. RUCH: Wypychanie bioder. UWAGA: To wersja zaawansowana, wymaga idealnej stabilności miednicy.", "czas_min": 3, "parametry": "3x10", "miesnie": "Pośladki"},
        {"nazwa": "Wykroki boczne (Side Lunge)", "opis": "POZYCJA: Stojąc, szeroki krok w bok. RUCH: Przenieś ciężar na nogę wykroczną, wypychając biodro w tył. UWAGA: Noga, która zostaje w miejscu, jest wyprostowana.", "czas_min": 3, "parametry": "3x12", "miesnie": "Pośladki, uda"},
        {"nazwa": "Mostki biodrowe na 1 nodze", "opis": "POZYCJA: Leżenie tyłem, jedna noga uniesiona. RUCH: Unoś biodra, używając tylko jednej nogi. UWAGA: Utrzymuj miednicę w poziomie przez cały ruch.", "czas_min": 2, "parametry": "3x12", "miesnie": "Pośladki"},
        {"nazwa": "Kettlebell Swing", "opis": "POZYCJA: Stojąc, kettlebell między nogami. RUCH: Wykonaj dynamiczny wymach biodrami. UWAGA: Plecy proste, ruch inicjują biodra, nie ręce.", "czas_min": 3, "parametry": "3x20", "miesnie": "Pośladki, dwugłowe"},
        {"nazwa": "Odwodzenie w leżeniu na boku", "opis": "POZYCJA: Leżenie na boku, noga prosta. RUCH: Unoś prostą nogę w górę. UWAGA: Palce stopy skierowane do przodu, nie do sufitu.", "czas_min": 2, "parametry": "3x20", "miesnie": "Pośladkowy średni"},
        {"nazwa": "Good Morning", "opis": "POZYCJA: Sztanga na karku, stopy na szerokość bioder. RUCH: Skłon w przód z wypchnięciem bioder w tył. UWAGA: Tylko do poziomu równoległego do podłogi.", "czas_min": 3, "parametry": "3x10", "miesnie": "Pośladki, prostowniki"},
        {"nazwa": "Przysiad w tempie (pauza)", "opis": "POZYCJA: Klasyczna. RUCH: Zejdź w dół (3 sekundy), zatrzymaj na 2 sekundy, wróć szybko. UWAGA: Stałe napięcie mięśniowe w pauzie.", "czas_min": 3, "parametry": "3x10", "miesnie": "Pośladki"},
        {"nazwa": "Glute Machine (maszyna)", "opis": "POZYCJA: Stojąc lub w klęku na maszynie. RUCH: Wypychanie nogi w tył przeciwko oporowi. UWAGA: Nie pogłębiaj lędźwi podczas ruchu.", "czas_min": 3, "parametry": "3x12", "miesnie": "Pośladkowy wielki"},
        {"nazwa": "Spacer kraba z gumą", "opis": "POZYCJA: Pozycja półprzysiadu, guma nad kolanami. RUCH: Krok w bok, utrzymując napięcie taśmy. UWAGA: Nie prostuj nóg w trakcie chodzenia.", "czas_min": 2, "parametry": "3x20m", "miesnie": "Pośladkowy średni"},
        {"nazwa": "Unoszenie nogi w klęku (z gumą)", "opis": "POZYCJA: Klęk, guma pod stopą i dłonią. RUCH: Unoś nogę do boku/tyłu. UWAGA: Stabilny tułów.", "czas_min": 2, "parametry": "3x15", "miesnie": "Pośladki"}
    ],
    "Zakończenie treningu": [
        {"nazwa": "Schładzanie na rowerku stacjonarnym", "opis": "POZYCJA: Siedząc na rowerku. RUCH: Spokojne pedałowanie z bardzo niskim oporem. UWAGA: Skup się na głębokim, wolnym oddechu, aby uspokoić tętno.", "czas_min": 5, "parametry": "5 minut", "miesnie": "Całe ciało, układ krążenia"},
        {"nazwa": "Rozciąganie taśmy tylnej w siadzie", "opis": "POZYCJA: Siad prosty, nogi złączone. RUCH: Wykonaj skłon w przód, próbując chwycić stopy. UWAGA: Nie szarp, pogłębiaj zakres z każdym wydechem.", "czas_min": 2, "parametry": "2 minuty", "miesnie": "Grupa kulszowo-goleniowa, łydki"},
        {"nazwa": "Pozycja dziecka (Child's Pose)", "opis": "POZYCJA: Klęk na podłodze, pośladki na piętach, ręce wyciągnięte w przód. RUCH: Opuść czoło do maty. UWAGA: Rozluźnij mięśnie grzbietu i poczuj wydłużanie kręgosłupa.", "czas_min": 2, "parametry": "2 minuty", "miesnie": "Prostowniki grzbietu, barki"},
        {"nazwa": "Rozciąganie klatki piersiowej przy ścianie", "opis": "POZYCJA: Stań bokiem do ściany, oprzyj przedramię na niej. RUCH: Skręć tułów w przeciwną stronę. UWAGA: Poczuj rozciąganie w obrębie klatki i przedniej części barku.", "czas_min": 2, "parametry": "1 minuta na stronę", "miesnie": "Klatka piersiowa, barki"},
        {"nazwa": "Pozycja Sfinksa", "opis": "POZYCJA: Leżenie przodem, podparcie na przedramionach. RUCH: Wypychaj klatkę w przód i w górę. UWAGA: Jeśli czujesz ból w lędźwiach, obniż tułów – to nie jest ćwiczenie siłowe.", "czas_min": 2, "parametry": "2 minuty", "miesnie": "Brzuch, prostowniki grzbietu"},
        {"nazwa": "Rozciąganie zginaczy bioder", "opis": "POZYCJA: Wykrok, tylne kolano na macie. RUCH: Wypchnij biodra w przód. UWAGA: Napnij pośladek nogi zakrocznej, aby zwiększyć rozciąganie z przodu uda.", "czas_min": 2, "parametry": "1 minuta na stronę", "miesnie": "Zginacze bioder, czworogłowe"},
        {"nazwa": "Kot-krowa (rozluźnienie)", "opis": "POZYCJA: Klęk podparty. RUCH: Naprzemienne wyginanie kręgosłupa w górę i w dół. UWAGA: Ruch w tempie oddechu – wydech w wygięciu, wdech w obniżeniu.", "czas_min": 2, "parametry": "2 minuty", "miesnie": "Cały kręgosłup"},
        {"nazwa": "Rozciąganie pośladków (Gołąb)", "opis": "POZYCJA: Siad ze zgiętą nogą przed sobą. RUCH: Pochyl tułów nad nogą. UWAGA: Utrzymuj biodra w linii, nie przechylaj się na boki.", "czas_min": 2, "parametry": "1 minuta na stronę", "miesnie": "Pośladki, rotator biodra"},
        {"nazwa": "Skręt tułowia w leżeniu", "opis": "POZYCJA: Leżenie tyłem, ręce szeroko. RUCH: Przenieś zgięte kolana w bok, dotykając nimi podłogi. UWAGA: Barki muszą przylegać do maty.", "czas_min": 2, "parametry": "1 minuta na stronę", "miesnie": "Mięśnie skośne brzucha, plecy"},
        {"nazwa": "Rozciąganie łydki przy ścianie", "opis": "POZYCJA: Stanie przy ścianie, stopa w tyle. RUCH: Dociśnij piętę do ziemi. UWAGA: Kolano nogi zakrocznej powinno być wyprostowane.", "czas_min": 2, "parametry": "1 minuta na stronę", "miesnie": "Brzuchaty łydki, płaszczkowaty"},
        {"nazwa": "Pozycja 'Pies z głową w dół'", "opis": "POZYCJA: Podpór, wypchnij biodra wysoko w górę. RUCH: Staraj się dociągnąć pięty do ziemi. UWAGA: Plecy muszą być proste.", "czas_min": 2, "parametry": "2 minuty", "miesnie": "Tył nóg, plecy"},
        {"nazwa": "Rozciąganie bicepsa przy ścianie", "opis": "POZYCJA: Dłoń oparta o ścianę, kciuk w górę. RUCH: Obróć tułów w przeciwną stronę. UWAGA: Poczuj rozciąganie wzdłuż ramienia.", "czas_min": 2, "parametry": "1 minuta na stronę", "miesnie": "Biceps, przedramiona"},
        {"nazwa": "Rozciąganie tricepsa za głową", "opis": "POZYCJA: Stojąc, zegnij rękę w łokciu. RUCH: Drugą ręką dociśnij łokieć do dołu. UWAGA: Nie wyginaj kręgosłupa, trzymaj brzuch napięty.", "czas_min": 2, "parametry": "1 minuta na stronę", "miesnie": "Triceps"},
        {"nazwa": "Wydłużanie szyi (boczne)", "opis": "POZYCJA: Siedząc prosto. RUCH: Przyciągnij ucho do barku, przeciwną rękę wyciągnij w dół. UWAGA: Nie szarp głową.", "czas_min": 2, "parametry": "1 minuta na stronę", "miesnie": "Mięśnie szyi, trapezius"},
        {"nazwa": "Pozycja 'Motyla' (stopy razem)", "opis": "POZYCJA: Siad, podeszwy stóp stykają się. RUCH: Delikatnie dociskaj kolana do podłogi. UWAGA: Plecy wyprostowane.", "czas_min": 2, "parametry": "2 minuty", "miesnie": "Przywodziciele ud"},
        {"nazwa": "Rozciąganie powięzi czworogłowego", "opis": "POZYCJA: Klęk na jednej nodze, pięta do pośladka. RUCH: Przyciągnij piętę ręką. UWAGA: Napnij pośladek, aby chronić lędźwie.", "czas_min": 2, "parametry": "1 minuta na stronę", "miesnie": "Czworogłowe"},
        {"nazwa": "Pozycja trupa (Savasana)", "opis": "POZYCJA: Leżenie tyłem, ręce wzdłuż ciała. RUCH: Pełne rozluźnienie. UWAGA: Obserwuj oddech przez 3 minuty, nic więcej nie rób.", "czas_min": 3, "parametry": "3 minuty", "miesnie": "Całe ciało (relaksacja)"},
        {"nazwa": "Krążenia nadgarstków i stawów skokowych", "opis": "POZYCJA: Siad lub stanie. RUCH: Powolne, obszerne krążenia stawami. UWAGA: Jeśli czujesz 'strzelanie', wykonuj ruch jeszcze wolniej.", "czas_min": 2, "parametry": "2 minuty", "miesnie": "Stawy obwodowe"},
        {"nazwa": "Rozciąganie najszerszego grzbietu (chwyt za słup)", "opis": "POZYCJA: Chwyć słup/framugę drzwi, odchyl biodra w tył. RUCH: Poczuj rozciąganie w boku pleców. UWAGA: Nie ciągnij na siłę.", "czas_min": 2, "parametry": "1 minuta na stronę", "miesnie": "Najszerszy grzbietu"},
        {"nazwa": "Głębokie oddychanie przeponowe", "opis": "POZYCJA: Leżenie, dłonie na brzuchu. RUCH: Wdech przez nos (brzuch rośnie), wydech przez usta. UWAGA: Wydech powinien być dwa razy dłuższy niż wdech.", "czas_min": 3, "parametry": "3 minuty", "miesnie": "Przepona"}
    ]
}

PLIK_WLASNYCH_CWICZEN = "wlasne_cwiczenia.json"

def zaladuj_wlasne_cwiczenia():
    if os.path.exists(PLIK_WLASNYCH_CWICZEN):
        try:
            with open(PLIK_WLASNYCH_CWICZEN, "r") as f:
                dane = json.load(f)
                for kat, lista in dane.get("FIZJO", {}).items():
                    if kat in BAZA_FIZJO: BAZA_FIZJO[kat].extend(lista)
                for kat, lista in dane.get("GYM", {}).items():
                    if kat in BAZA_SILOWNIA: BAZA_SILOWNIA[kat].extend(lista)
        except: pass

zaladuj_wlasne_cwiczenia()
GLOBALNA_BAZA = {**BAZA_FIZJO, **BAZA_SILOWNIA}

# AUTOMATYCZNE SORTOWANIE ALFABETYCZNE BAZ
for compliance_baza in [BAZA_FIZJO, BAZA_SILOWNIA, GLOBALNA_BAZA]:
    for compliance_kat in compliance_baza:
        compliance_baza[compliance_kat] = sorted(compliance_baza[compliance_kat], key=lambda compliance_x: compliance_x['nazwa'].lower())

# ==============================================================================
# LOGIKA GENERATORÓW
# ==============================================================================
def generuj_plan(profil, budzet, dni):
    plan = []
    b_fizjo = {k: list(v) for k, v in BAZA_FIZJO.items()}
    b_gym = {k: list(v) for k, v in BAZA_SILOWNIA.items()}

    def pop_random(baza, kat):
        if baza.get(kat) and len(baza[kat]) > 0:
            idx = random.randint(0, len(baza[kat])-1)
            cw = baza[kat].pop(idx).copy()
            cw["uwagi"] = ""
            return cw
        return None

    czy_split = "Split" in profil
    czy_wielo_dniowy = czy_split or "Kompleksowy" in profil
    rzeczywista_liczba_dni = dni if czy_wielo_dniowy else 1
    is_gym = "GYM:" in profil

    dni_tygodnia = ["Poniedziałek", "Wtorek", "Środa", "Czwartek", "Piątek", "Sobota", "Niedziela"]
    
    if czy_split:
        if is_gym:
            uklady = {
                1: [["Klatka piersiowa", "Plecy", "Nogi", "Ręce", "Pośladki"]],
                2: [["Klatka piersiowa", "Ręce"], ["Plecy", "Nogi", "Pośladki"]],
                3: [["Klatka piersiowa", "Ręce"], ["Nogi", "Pośladki"], ["Plecy"]],
                4: [["Ręce"], ["Klatka piersiowa"], ["Nogi"], ["Plecy"]],
                5: [["Klatka piersiowa"], ["Plecy"], ["Nogi"], ["Ręce"], ["Pośladki"]],
                6: [["Klatka piersiowa"], ["Plecy"], ["Nogi"], ["Ręce"], ["Pośladki"], ["Klatka piersiowa", "Plecy"]],
                7: [["Klatka piersiowa"], ["Plecy"], ["Nogi"], ["Ręce"], ["Pośladki"], ["Klatka piersiowa"], ["Plecy", "Nogi"]]
            }
        else:
            uklady = {
                1: [["Głowa/Szyja", "Kończyna górna", "Core (Tułów)", "Kończyna dolna"]],
                2: [["Kończyna górna", "Głowa/Szyja"], ["Kończyna dolna", "Core (Tułów)"]],
                3: [["Kończyna górna", "Głowa/Szyja"], ["Core (Tułów)"], ["Kończyna dolna"]],
                4: [["Głowa/Szyja"], ["Kończyna górna"], ["Core (Tułów)"], ["Kończyna dolna"]],
                5: [["Głowa/Szyja"], ["Kończyna górna"], ["Core (Tułów)"], ["Kończyna dolna"], ["Core (Tułów)", "Kończyna górna"]],
                6: [["Głowa/Szyja"], ["Kończyna górna"], ["Core (Tułów)"], ["Kończyna dolna"], ["Kończyna górna"], ["Core (Tułów)"]],
                7: [["Głowa/Szyja"], ["Kończyna górna"], ["Core (Tułów)"], ["Kończyna dolna"], ["Kończyna górna"], ["Core (Tułów)"], ["Kończyna dolna"]]
            }
        plan_na_dni = uklady.get(rzeczywista_liczba_dni, uklady[7][:rzeczywista_liczba_dni])

    for i in range(rzeczywista_liczba_dni):
        if czy_wielo_dniowy:
            if czy_split:
                dzien = dni_tygodnia[i]
                partie = plan_na_dni[i]
                partie_str = ' + '.join([p.upper() for p in partie])
                nazwa_dnia = f"{dzien}: {partie_str}"
                naglowek = {"nazwa": nazwa_dnia, "typ": "Split", "partie": partie_str, "opis": "", "czas_min": 0, "parametry": "-", "miesnie": "-", "uwagi": ""}
            else:
                nazwa_dnia = f"Dzień {i+1} (FIZJO: Kompleksowy)"
                naglowek = {"nazwa": nazwa_dnia, "typ": "Kompleksowy", "partie": "", "opis": "", "czas_min": 0, "parametry": "-", "miesnie": "-", "uwagi": ""}
            plan.append(("NAGŁÓWEK DNIA", naglowek))

        if is_gym:
            cw_start = pop_random(b_gym, "Rozgrzewka") or random.choice(BAZA_SILOWNIA["Rozgrzewka"]).copy()
            cw_start["uwagi"] = ""
            plan.append(("GYM: Rozgrzewka", cw_start))
        else:
            cw_start = pop_random(b_fizjo, "Oddechowe") or random.choice(BAZA_FIZJO["Oddechowe"]).copy()
            cw_start["uwagi"] = ""
            plan.append(("Oddechowe (Rozgrzewka)", cw_start))
        
        realny_czas = cw_start.get("czas_min", 2) if not is_gym else 0

        if czy_split:
            partie = plan_na_dni[i]
            for p in partie:
                dodano = 0
                while dodano < budzet:
                    baza_docelowa = b_gym if is_gym else b_fizjo
                    cw = pop_random(baza_docelowa, p)
                    if not cw: break
                    etykieta = f"GYM: {p}" if is_gym else p
                    plan.append((etykieta, cw))
                    dodano += 1
        else:
            if is_gym:
                if "Ogólnorozwojowy" in profil:
                    for p in ["Klatka piersiowa", "Plecy", "Nogi", "Ręce", "Pośladki"]:
                        dodano = 0
                        while dodano < budzet:
                            cw = pop_random(b_gym, p)
                            if not cw: break
                            plan.append((f"GYM: {p}", cw))
                            dodano += 1
                else:
                    kat = profil.split(" - ")[1]
                    dodano = 0
                    while dodano < budzet:
                        cw = pop_random(b_gym, kat)
                        if not cw: break
                        plan.append((f"GYM: {kat}", cw))
                        dodano += 1
            else:
                czas_koncowy = 2 
                if "Kompleksowy" in profil:
                    lancuch = ["Głowa/Szyja", "Kończyna górna", "Core (Tułów)", "Kończyna dolna"]
                    puste = 0
                    while realny_czas + czas_koncowy < budzet and puste < len(lancuch):
                        dodano_cos = False
                        for kat in lancuch:
                            dostepne = [c for c in b_fizjo.get(kat, []) if realny_czas + c["czas_min"] + czas_koncowy <= budzet]
                            if dostepne:
                                cw_wybrane = dostepne[random.randint(0, len(dostepne)-1)]
                                cw = b_fizjo[kat].pop(b_fizjo[kat].index(cw_wybrane)).copy()
                                cw["uwagi"] = ""
                                plan.append((kat, cw))
                                realny_czas += cw["czas_min"]
                                dodano_cos = True
                        if not dodano_cos: puste += 1
                else:
                    kat = profil.split(" - ")[1].replace("Tylko ", "")
                    while realny_czas + czas_koncowy < budzet:
                        dostepne = [c for c in b_fizjo.get(kat, []) if realny_czas + c["czas_min"] + czas_koncowy <= budzet]
                        if not dostepne: break
                        cw_wybrane = dostepne[random.randint(0, len(dostepne)-1)]
                        cw = b_fizjo[kat].pop(b_fizjo[kat].index(cw_wybrane)).copy()
                        cw["uwagi"] = ""
                        plan.append((kat, cw))
                        realny_czas += cw["czas_min"]

        if is_gym:
            cw_koniec = pop_random(b_gym, "Zakończenie treningu") or random.choice(BAZA_SILOWNIA["Zakończenie treningu"]).copy()
            cw_koniec["uwagi"] = ""
            plan.append(("GYM: Zakończenie", cw_koniec))
        else:
            cw_koniec = pop_random(b_fizjo, "Oddechowe") or random.choice(BAZA_FIZJO["Oddechowe"]).copy()
            cw_koniec["uwagi"] = ""
            plan.append(("Oddechowe (Wyciszenie)", cw_koniec))

    st.session_state.wylosowany_plan_cache = plan
    st.session_state.is_gym = is_gym

# ==============================================================================
# REJESTR EKSPORTÓW
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
    if not st.session_state.wylosowany_plan_cache: 
        return None
    
    dane_excel = []
    is_gym = st.session_state.get('is_gym', False)
    
    dane_excel.append(["Imię i Nazwisko:", "", "", "", "", "", ""])
    dane_excel.append(["Płeć:", "", "", "", "", "", ""])
    
    realny_czas = sum(compliance_x[1].get('czas_min', 0) for compliance_x in st.session_state.wylosowany_plan_cache if compliance_x[0] != "NAGŁÓWEK DNIA")
    if is_gym:
        dane_excel.append(["Podsumowanie planu:", "Trening Siłowy (Ilościowy)", "", "", "", "", ""])
    else:
        dane_excel.append(["Całkowity czas planu:", f"{realny_czas} min", "", "", "", "", ""])
        
    dane_excel.append(["", "", "", "", "", "", ""])
    czy_wielodniowy_excel = any(kat == "NAGŁÓWEK DNIA" for kat, cw in st.session_state.wylosowany_plan_cache)
    
    if czy_wielodniowy_excel:
        czasy_dni = []
        for k, c in st.session_state.wylosowany_plan_cache:
            if k == "NAGŁÓWEK DNIA": czasy_dni.append(0)
            else:
                if czasy_dni: czasy_dni[-1] += c.get('czas_min', 0)

        lp = 1
        dzien_aktualny = 0
        for idx, (kat, cw) in enumerate(st.session_state.wylosowany_plan_cache):
            if kat == "NAGŁÓWEK DNIA":
                if dzien_aktualny > 0: dane_excel.append(["", "", "", "", "", "", ""])
                dzien_aktualny += 1
                nazwa_dnia_i_partii = cw['nazwa']
                if is_gym: dane_excel.append([f"{nazwa_dnia_i_partii}", "", "", "", "", "", ""])
                else:
                    czas_tego_dnia = czasy_dni[dzien_aktualny - 1] if dzien_aktualny <= len(czasy_dni) else 0
                    dane_excel.append([f"{nazwa_dnia_i_partii} - Czas trwania: {czas_tego_dnia} min", "", "", "", "", "", ""])
                dane_excel.append(["L.p.", "Nazwa ćwiczenia", "Czas", "Ilość serii", "Ilość powtórzeń", "Sposób wykonania", "Uwagi"])
                lp = 1
            else:
                p_lower = str(cw['parametry']).lower()
                if 'x' in p_lower:
                    parts = p_lower.split('x')
                    serie = parts[0].replace('serie','').replace('seria','').strip()
                    powt = parts[1].replace('powtórzeń','').replace('powtórzenia','').strip()
                else:
                    serie = "1"
                    powt = cw['parametry'].strip()
                czas_str = "-" if is_gym else f"{cw.get('czas_min', 0)} min"
                dane_excel.append([lp, cw['nazwa'], czas_str, serie, powt, cw['opis'], cw.get('uwagi', '')])
                lp += 1
    else:
        for dzien in range(1, liczba_dni + 1):
            dane_excel.append([f"DZIEŃ {dzien}", "", "", "", "", "", ""])
            dane_excel.append(["L.p.", "Nazwa ćwiczenia", "Czas", "Ilość serii", "Ilość powtórzeń", "Sposób wykonania", "Uwagi"])
            for compliance_idx, (kat, cw) in enumerate(st.session_state.wylosowany_plan_cache, 1):
                p_lower = str(cw['parametry']).lower()
                if 'x' in p_lower:
                    parts = p_lower.split('x')
                    serie = parts[0].replace('serie','').replace('seria','').strip()
                    powt = parts[1].replace('powtórzeń','').replace('powtórzenia','').strip()
                else:
                    serie = "1"
                    powt = cw['parametry'].strip()
                czas_str = "-" if is_gym else f"{cw.get('czas_min', 0)} min"
                dane_excel.append([compliance_idx, cw['nazwa'], czas_str, serie, powt, cw['opis'], cw.get('uwagi', '')])

    df = pd.DataFrame(dane_excel)
    bio = io.BytesIO()
    writer = pd.ExcelWriter(bio, engine='openpyxl')
    df.to_excel(writer, index=False, header=False, sheet_name='Harmonogram')
    worksheet = writer.sheets['Harmonogram']
    
    for c_letter, c_width in [('A', 8), ('B', 30), ('C', 10), ('D', 12), ('E', 20), ('F', 50), ('G', 25)]:
        worksheet.column_dimensions[c_letter].width = c_width

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    fill_day = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_header = PatternFill(start_color="8DB4E2", end_color="8DB4E2", fill_type="solid")
    fill_data_even = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    
    data_row_counter = 0
    for r_idx, r_data in enumerate(dane_excel, 1):
        val_a = str(r_data[0])
        if r_idx in [1, 2, 3]:
            worksheet.cell(row=r_idx, column=1).font = Font(bold=True)
        elif any(val_a.startswith(prefix) for prefix in ["DZIEŃ", "Poniedziałek", "Wtorek", "Środa", "Czwartek", "Piątek", "Sobota", "Niedziela", "Dzień"]):
            worksheet.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=7)
            cell = worksheet.cell(row=r_idx, column=1)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = fill_day
            cell.alignment = Alignment(horizontal='center', vertical='center')
            for c_idx in range(1, 8): worksheet.cell(row=r_idx, column=c_idx).border = thin_border
            data_row_counter = 0
        elif val_a == "L.p.":
            for c_idx in range(1, 8):
                cell = worksheet.cell(row=r_idx, column=c_idx)
                cell.font = Font(bold=True)
                cell.fill = fill_header
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
        elif val_a.isdigit():
            data_row_counter += 1
            current_fill = fill_data_even if data_row_counter % 2 == 0 else None
            for c_idx in range(1, 8):
                cell = worksheet.cell(row=r_idx, column=c_idx)
                cell.border = thin_border
                if current_fill: cell.fill = current_fill
                cell.alignment = Alignment(wrap_text=True, vertical='top') if c_idx in [6, 7] else Alignment(horizontal='center', vertical='center')
    writer.close()
    return bio.getvalue()

# ==============================================================================
# UI - INTERFEJS STRONY WEBOWEJ
# ==============================================================================
st.title("Fizjo Workout Ultimate")
st.markdown("Zintegrowane środowisko projektowania programów treningowych.")

with st.sidebar:
    st.header("🔑 Dostęp do AI")
    user_api_key = st.text_input("Twój klucz API Groq:", type="password", help="Pobierz darmowy klucz ze strony console.groq.com")
    st.caption("Nie masz klucza? Pobierz go bezpłatnie ze strony: [console.groq.com](https://console.groq.com/)")
    groq_client = Groq(api_key=user_api_key) if user_api_key else None

    st.divider()
    st.header("⚙️ Konfiguracja")
    
    st.markdown("""
    <style>
    div[role="radiogroup"] label {
        padding: 6px 10px;
        border-radius: 8px;
        margin-bottom: 4px;
        transition: 0.3s;
    }
    div[role="radiogroup"] label:has(div:contains("FIZJO")) {
        background-color: rgba(40, 167, 69, 0.15);
        border-left: 4px solid #28a745;
    }
    div[role="radiogroup"] label:has(div:contains("GYM")) {
        background-color: rgba(0, 123, 255, 0.15);
        border-left: 4px solid #007bff;
    }
    </style>
    """, unsafe_allow_html=True)
    
    profil = st.radio("Profil Silnika:", [
        "🟩 FIZJO: Kompleksowy (Wszystkie partie)", 
        "🟩 FIZJO: Automatyczny Split (Dni Tygodnia)",
        "🟩 FIZJO: Ukierunkowany - Tylko Oddechowe",
        "🟩 FIZJO: Ukierunkowany - Tylko Głowa/Szyja",
        "🟩 FIZJO: Ukierunkowany - Tylko Kończyna górna",
        "🟩 FIZJO: Ukierunkowany - Tylko Core (Tułów)",
        "🟩 FIZJO: Ukierunkowany - Tylko Kończyna dolna",
        "🟦 GYM: Ogólnorozwojowy (FBW - Całe Ciało)",
        "🟦 GYM: Automatyczny Split (Dni Tygodnia)",
        "🟦 GYM: Ukierunkowany - Klatka piersiowa",
        "🟦 GYM: Ukierunkowany - Ręce",
        "🟦 GYM: Ukierunkowany - Plecy",
        "🟦 GYM: Ukierunkowany - Nogi",
        "🟦 GYM: Ukierunkowany - Pośladki"
    ])
    
    is_gym = "GYM:" in profil
    czy_split = "Split" in profil
    czy_wielo_dniowy = czy_split or "Kompleksowy" in profil
    
    budzet = st.number_input("Ilość ćw. NA PARTIĘ:" if is_gym else "Budżet czasu (min):", min_value=1, max_value=120, value=4 if is_gym else 45)
    dni = st.number_input("Liczba dni (tylko dla Split/Kompleksowy):", min_value=1, max_value=7, value=4, disabled=not czy_wielo_dniowy)
    
    if st.button("⚡ GENERUJ AUTOMAT", use_container_width=True, type="primary"):
        generuj_plan(profil, budzet, dni)
        st.rerun()
        
    if st.button("❌ CZYŚĆ EKRAN (RESET)", use_container_width=True):
        st.session_state.wylosowany_plan_cache = []
        st.rerun()
        
    st.divider()
    if st.session_state.wylosowany_plan_cache:
        st.success("Plan gotowy!")
        st.download_button("💾 POBIERZ DOCX", generuj_docx(), "Plan_Treningowy.docx", use_container_width=True)
        # PRZYWRÓCONY PRZYCISK POBIERANIA EXCELA
        st.download_button("📊 POBIERZ EXCEL", generuj_excel(dni), "Plan_Treningowy.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# ZAKŁADKI GŁÓWNE
tab1, tab2, tab3, tab4 = st.tabs(["📝 Twój Plan", "➕ Kreator", "✨ Asystent AI", "⚙️ Baza Ćwiczeń"])

# ZAKŁADKA 1: WYGENEROWANY PLAN
with tab1:
    if not st.session_state.wylosowany_plan_cache:
        st.info("👈 Użyj panelu bocznego, aby wygenerować plan lub przejdź do Kreatora.")
    else:
        def odswiez_nazwy_dni(blocks):
            dni_tyg = ["Poniedziałek", "Wtorek", "Środa", "Czwartek", "Piątek", "Sobota", "Niedziela"]
            for compliance_i, block in enumerate(blocks):
                kat, cw = block[0]
                if kat == "NAGŁÓWEK DNIA":
                    typ = cw.get("typ", "")
                    if typ == "Split":
                        nowy_dzien = dni_tyg[compliance_i] if compliance_i < 7 else f"Dzień {compliance_i+1}"
                        cw["nazwa"] = f"{nowy_dzien}: {cw['partie']}"
                    elif typ == "Kompleksowy":
                        cw["nazwa"] = f"Dzień {compliance_i+1} (FIZJO: Kompleksowy)"

        days_blocks = []
        current_block = []
        for item in st.session_state.wylosowany_plan_cache:
            if item[0] == "NAGŁÓWEK DNIA":
                if current_block: days_blocks.append(current_block)
                current_block = [item]
            else: current_block.append(item)
        if current_block: days_blocks.append(current_block)

        abs_idx = 0
        licznik = 1
        
        for block_idx, block in enumerate(days_blocks):
            for item_idx_in_block, (kat, cw) in enumerate(block):
                if kat == "NAGŁÓWEK DNIA":
                    st.markdown(f"### 📅 {cw['nazwa']}")
                    c1, c2, c3, c4 = st.columns([1.5, 1.5, 2, 5])
                    if block_idx > 0:
                        if c1.button("⬆️ Dzień", key=f"up_day_{abs_idx}", use_container_width=True):
                            days_blocks[block_idx], days_blocks[block_idx-1] = days_blocks[block_idx-1], days_blocks[block_idx]
                            odswiez_nazwy_dni(days_blocks)
                            st.session_state.wylosowany_plan_cache = [i for b in days_blocks for i in b]
                            st.rerun()
                    if block_idx < len(days_blocks) - 1:
                        if c2.button("⬇️ Dzień", key=f"down_day_{abs_idx}", use_container_width=True):
                            days_blocks[block_idx], days_blocks[block_idx+1] = days_blocks[block_idx+1], days_blocks[block_idx]
                            odswiez_nazwy_dni(days_blocks)
                            st.session_state.wylosowany_plan_cache = [i for b in days_blocks for i in b]
                            st.rerun()
                    if c3.button("❌ Usuń cały", key=f"del_day_{abs_idx}", type="primary", use_container_width=True):
                        days_blocks.pop(block_idx)
                        odswiez_nazwy_dni(days_blocks)
                        st.session_state.wylosowany_plan_cache = [i for b in days_blocks for i in b]
                        st.rerun()
                    licznik = 1
                    abs_idx += 1
                    continue
                    
                with st.expander(f"{licznik}. {cw['nazwa']} ({kat})", expanded=True):
                    col1, col2 = st.columns([3, 2])
                    with col1:
                        nowe_parametry = st.text_input("Zalecenie", cw['parametry'], key=f"p_{abs_idx}")
                        nowe_uwagi = st.text_input("Uwagi", cw.get('uwagi', ''), key=f"u_{abs_idx}")
                        st.session_state.wylosowany_plan_cache[abs_idx][1]['parametry'] = nowe_parametry
                        st.session_state.wylosowany_plan_cache[abs_idx][1]['uwagi'] = nowe_uwagi
                        st.caption(f"**Anatomia:** {cw['miesnie']}")
                        st.write(cw['opis'])
                    with col2:
                        c_up, c_down, c_del = st.columns(3)
                        moze_w_gore = abs_idx > 0 and st.session_state.wylosowany_plan_cache[abs_idx-1][0] != "NAGŁÓWEK DNIA"
                        moze_w_dol = abs_idx < len(st.session_state.wylosowany_plan_cache) - 1 and st.session_state.wylosowany_plan_cache[abs_idx+1][0] != "NAGŁÓWEK DNIA"
                        if moze_w_gore and c_up.button("⬆️", key=f"up_{abs_idx}"):
                            st.session_state.wylosowany_plan_cache[abs_idx], st.session_state.wylosowany_plan_cache[abs_idx-1] = st.session_state.wylosowany_plan_cache[abs_idx-1], st.session_state.wylosowany_plan_cache[abs_idx]
                            st.rerun()
                        if moze_w_dol and c_down.button("⬇️", key=f"down_{abs_idx}"):
                            st.session_state.wylosowany_plan_cache[abs_idx], st.session_state.wylosowany_plan_cache[abs_idx+1] = st.session_state.wylosowany_plan_cache[abs_idx+1], st.session_state.wylosowany_plan_cache[abs_idx]
                            st.rerun()
                        if c_del.button("❌", key=f"del_{abs_idx}", type="primary"):
                            st.session_state.wylosowany_plan_cache.pop(abs_idx)
                            st.rerun()
                licznik += 1
                abs_idx += 1

# ZAKŁADKA 2: KREATOR MANUALNY
with tab2:
    st.subheader("Manualne dodawanie ćwiczeń")
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
    st.subheader("Wirtualny Konsultant Treningowy")
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
