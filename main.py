import json
import os
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading # <--- NOWY IMPORT (do płynnego działania czatu)

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

import customtkinter as ctk
from PIL import Image

from groq import Groq

# Inicjalizacja klienta Groq. Wklej tutaj swój klucz API.
GROQ_API_KEY = "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"
groq_client = Groq(api_key=GROQ_API_KEY)

# Konfiguracja głównego motywu CustomTkinter
# ... (reszta Twojego kodu bez zmian)

# ==============================================================================
# KONFIGURACJA GŁÓWNEGO MOTYWU UI
# ==============================================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

# ==============================================================================
# 1. AUTORSKA BAZA FIZJOTERAPEUTYCZNA (ODDECH, SZYJA, CORE, KOŃCZYNY)
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
        {"nazwa": "Głęboki wdech w pozycji embrionalnej", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Ukłon japoński (siad na piętach, tułów na udach, ręce wyciągnięte). Głębokie wdechy rozszerzające tylną ścianę klatki piersiowej.", "czas_min": 3, "parametry": "3 minuty", "miesnie": "Mięśnie międzyżebrowe tylne, przepona"}
    ],
    "Głowa/Szyja": [
        {"nazwa": "Retrakcja szyi (Cofanie brody)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: W pozycji siedzącej cofnij głowę w płaszczyźnie poziomej (zrób 'podwójny podbródek'). Wzrok prosto, trzymaj 3 s.", "czas_min": 2, "parametry": "3 serie x 10 powtórzeń", "miesnie": "Mięśnie głębokie zginacze szyi, mięsień płatowaty głowy"},
        {"nazwa": "Izometryczne parcie w przód", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Przyłóż dłoń do czoła. Naciskaj głową w przód na stawiającą opór dłoń, nie pozwalając na ruch. Trzymaj przez 5 sekund.", "czas_min": 2, "parametry": "3 serie x 5s trzymania", "miesnie": "Mięsień mostkowo-obojczykowo-sutkowy (MOS), zginacze długie"},
        {"nazwa": "Izometryczne parcie w bok", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Przyłóż dłoń nad uchem. Naciskaj głową w bok przeciwko oporowi ręki. Trzymaj 5 sekund, zmień strony.", "czas_min": 2, "parametry": "3 serie x 5s na stronę", "miesnie": "Mięśnie pochyłe, mięsień płatowaty szyi"},
        {"nazwa": "Izometryczne parcie w tył", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Spleć dłonie z tyłu głowy. Naciskaj tyłem głowy w dłonie, utrzymując wzrok prosto. Utrzymaj 5 sekund.", "czas_min": 2, "parametry": "3 serie x 5s trzymania", "miesnie": "Mięsień podpotyliczny, prostownik grzbietu szyjny"},
        {"nazwa": "Rozciąganie trapeziusa", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Siedząc, opuść jedno ramię w dół. Drugą ręką delikatnie przyciągaj głowę DO przeciwnego barku.", "czas_min": 2, "parametry": "2 serie x 30 sekund", "miesnie": "Mięsień czworoboczny (część górna), dźwigacz łopatki"},
        {"nazwa": "Rozciąganie dźwigacza łopatki", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Skręć głowę o 45 stopni w bok, a następnie skieruj brodę w dół do pachy. Delikatnie wspomóż ruch ręką.", "czas_min": 2, "parametry": "2 serie x 30 sekund na stronę", "miesnie": "Mięsień dźwigacz łopatki, mięśnie podpotyliczne"},
        {"nazwa": "Otwieranie ust z oporem", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Podeprzyj żuchwę od dołu palcami. Otwieraj usta powoli, pokonując delikatny, stały opór stawiany przez dłoń.", "czas_min": 2, "parametry": "3 serie x 10 powtórzeń", "miesnie": "Mięśnie nadgnykowe, skrzydłowe boczne"},
        {"nazwa": "Ruchy sakadyczne oczu", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Trzymaj głowę nieruchomo. Przenoś wzrok szybko i precyzyjnie między dwoma punktami skrajnie po lewej i prawej.", "czas_min": 1, "parametry": "1 minuta", "miesnie": "Mięśnie gałkoruchowe, stymulacja układu przedsionkowego"},
        {"nazwa": "Automobilizacja z wałkiem", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie tyłem, twardy wałek pod podstawą czaszki. Wykonuj małe, powolne ruchy potakiwania oraz przeczenia.", "czas_min": 3, "parametry": "3 minuty", "miesnie": "Mięśnie podpotyliczne, powięź czepca ścięgnistego"},
        {"nazwa": "Retrakcja z rotacją w leżeniu", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie przodem, czoło nad podłogą. Wykonuj retrakcję szyi, a następnie delikatny obrót głowy w prawo i w lewo.", "czas_min": 2, "parametry": "10 powtórzeń na stronę", "miesnie": "Mięsień długi szyi, rotator odcinka szyjnego"}
    ],
    "Kończyna górna": [
        {"nazwa": "Wznosy ramion w pozycji Y", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie przodem, ramiona pod kątem 45 stopni (litera Y), kciuki w górę. Unoś ramiona, łącząc dolne kąty łopatek.", "czas_min": 3, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Mięsień czworoboczny (część dolna), mięsień równoległoboczny, aktony tyłu barku"},
        {"nazwa": "Rotacja zewnętrzna z gumą", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Łokcie ugięte 90 stopni przy bokach. Trzymając gumę, rozciągaj ją na boki poprzez rotację zewnętrzną w barkach.", "czas_min": 2, "parametry": "3 serie x 15 powtórzeń", "miesnie": "Mięsień podgrzebieniowy, mięsień obły mniejszy (stożek rotatorów)"},
        {"nazwa": "Pompki plus (Scapular)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Pozycja podporu przodem. Bez uginania łokci oddalaj łopatki od siebie (zaokrąglaj góry pleców) i zbliżaj.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Mięsień zębaty przedni, mięsień piersiowy mniejszy"},
        {"nazwa": "Rozciąganie zginaczy nadgarstka", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Wyprostuj rękę w łokciu przed sobą, dłoń w górę (palce w dół). Drugą ręką chwyć palce i przyciągaj je do siebie.", "czas_min": 2, "parametry": "2 serie x 30 sekund", "miesnie": "Mięsień zginacz promieniowy nadgarstka, zginacz powierzchowny palców"},
        {"nazwa": "Rozciąganie prostowników nadgarstka", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Wyprostuj rękę w łokciu, dłoń skierowana w dół. Drugą ręką dociśnij dłoń w dół, delikatnie zginając nadgarstek.", "czas_min": 2, "parametry": "2 serie x 30 sekund", "miesnie": "Mięsień prostownik wspólny palców, prostownik promieniowy długi"},
        {"nazwa": "Odwodzenie hantli w opadzie", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Opad tułowia, ręce wiszą luźno z hantlami. Unoś ramiona na boki do wysokości tułowia, angażując tył barku.", "czas_min": 3, "parametry": "3 serie x 10 powtórzeń", "miesnie": "Mięsień naramienny (akton tylny), mięsień równoległoboczny"},
        {"nazwa": "Ślizganie ramion po ścianie", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Oprzyj się plecami i głową o ścianę. Ręce ugięte 90 stopni. Przesuwaj ramiona pionowo w górę i w dół.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Mięsień czworoboczny dolny, rotator barku, zębaty przedni"},
        {"nazwa": "Uginanie przedramion z supinacją", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: W staniu, hantle chwytem neutralnym. Uginaj łokcie, rotując nadgarstki tak, by w górnej fazie dłonie patrzyły na Ciebie.", "czas_min": 2, "parametry": "3 serie x 15 powtórzeń", "miesnie": "Mięsień dwugłowy ramienia (biceps), m. ramienno-promieniowy"},
        {"nazwa": "Prostowanie przedramion w opadzie", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Opad tułowia, ramiona przy ciele, łonieć ugięty 90 stopni. Prostuj ramię w tył do pełnego wyprostu, blokując łokieć.", "czas_min": 3, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Mięsień trójgłowy ramienia (triceps), akton tylny naramiennego"},
        {"nazwa": "Spacer farmera z kettlebell", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Chwyć ciężkie odważniki w obie dłonie. Idź wyprostowany powolnym krokiem, utrzymując stabilne barki i napięty brzuch.", "czas_min": 3, "parametry": "3 serie x 45 sekund marszu", "miesnie": "Mięsień czworoboczny (góra), mięśnie przedramion (chwyt), core"}
    ],
    "Core (Tułów)": [
        {"nazwa": "Plank (Podpór przodem)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Oprzyj się na przedramionach i palcach stóp. Ciało w jednej linii, brzuch i pośladki mocno spięte, miednica neutralnie.", "czas_min": 3, "parametry": "3 serie x 30 sekund", "miesnie": "Mięsień poprzeczny brzucha, mięsień prosty brzucha, pośladkowy wielki"},
        {"nazwa": "Side Plank (Podpór bokiem)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie na boku, podparcie na przedramieniu i boku stopy. Unieś biodra w górę, tworząc linię prostą od głowy do stóp.", "czas_min": 3, "parametry": "2 serie x 20 sekund na stronę", "miesnie": "Mięsień czworoboczny lędźwi, mięśnie skośne brzucha, pośladkowy średni"},
        {"nazwa": "Dead Bug (Martwy robak)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie tyłem, ręce pionowo, nogi ugięte 90/90. Opuszczaj jednocześnie przeciwną rękę i nogę tuż nad podłogę, lędźwie w matę.", "czas_min": 3, "parametry": "3 serie x 10 powtórzeń na stronę", "miesnie": "Mięsień poprzeczny brzucha, mięsień skośny wewnętrzny"},
        {"nazwa": "Bird Dog (Pies-ptak)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Klęk podparty. Unieś i wyciągnij jednocześnie prawą rękę w przód i lewą nogę w tył do linii tułowia. Nie kołysz biodrami.", "czas_min": 3, "parametry": "3 serie x 10 powtórzeń na stronę", "miesnie": "Mięsień prostownik grzbietu, mięsień pośladkowy wielki, core"},
        {"nazwa": "Hollow Body (Pozycja kołyski)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie tyłem. Oderwij łopatki i proste nogi od podłogi. Odcinek lędźwiowy musi być maksymalnie dociśnięty do maty.", "czas_min": 2, "parametry": "3 serie x 20 sekund", "miesnie": "Mięsień prosty brzucha, mięsień biodrowo-lędźwiowy"},
        {"nazwa": "Bear Crawl Hold (Niedźwiedź)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Klęk podparty, palce stóp zaparte. Unieś kolana tylko 2-3 cm nad ziemię. Utrzymuj proste plecy i mocno napięty brzuch.", "czas_min": 2, "parametry": "3 serie x 20 sekund", "miesnie": "Mięsień poprzeczny brzucha, mięsień czworogłowy uda"},
        {"nazwa": "Niedźwiedź dynamiczny", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Z podporu niedźwiedzia wykonuj małe kroki w przód i w tył (naprzemienna praca ręki i przeciwnej nogi), kolana bardzo nisko.", "czas_min": 3, "parametry": "3 serie x 30 sekund", "miesnie": "Mięsień poprzeczny brzucha, mięśnie naramienne, m. głębokie uda"},
        {"nazwa": "Opuszczanie obustronne nóg", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie tyłem, nogi uniesione 90 stopni. Powoli opuszczaj obie złączone proste nogi, kontrolując, by lędźwie nie odrywały się.", "czas_min": 2, "parametry": "3 serie x 10 powtórzeń", "miesnie": "Mięsień prosty brzucha (dolna część), m. biodrowo-lędźwiowy"},
        {"nazwa": "Skośne spięcia brzucha", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie tyłem, kolana ugięte. Unosząc łopatki, kieruj prawe ramię w stronę lewego kolana. Na zmianę.", "czas_min": 2, "parametry": "3 serie x 15 powtórzeń", "miesnie": "Mięśnie skośne zewnętrzne i wewnętrzne brzucha"},
        {"nazwa": "Russian Twist", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Siad z uniesionymi stopami i odchylonym tułowiem. Rotuj klatkę piersiową i przekładaj dłonie z prawej na lewą stronę blisko bioder.", "czas_min": 2, "parametry": "3 serie x 20 skrętów", "miesnie": "Mięśnie skośne brzucha, mięsień poprzeczny brzucha"}
    ],
    "Kończyna dolna": [
        {"nazwa": "Przysiad klasyczny (Squat)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Stopy na szerokość barków. Schodź biodrami w dół i w tył, utrzymując kolana w linii stóp oraz zachowując proste plecy.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń", "miesnie": "Mięsień czworogłowy uda, mięsień pośladkowy wielki, przywodziciel wielki"},
        {"nazwa": "Mostki biodrowe (Glute Bridge)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie tyłem, kolana ugięte, stopy na ziemi. Unoś biodra w górę poprzez mocne spięcie pośladków, zablokuj na górze na 1 s.", "czas_min": 2, "parametry": "3 serie x 15 powtórzeń", "miesnie": "Mięsień pośladkowy wielki, mięsień dwugłowy uda, m. półbłoniasty"},
        {"nazwa": "Wznosy na palce stojąc", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Stojąc stabilnie, unoś pięty maksymalnie w górę, przechodząc na palce stóp, a następnie powoli i kontrolowanie opuszczaj.", "czas_min": 2, "parametry": "3 serie x 20 powtórzeń", "miesnie": "Mięsień brzuchaty łydki, mięsień płaszczkowaty"},
        {"nazwa": "Zakroki (Reverse Lunges)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Z pozycji stojącej wykonaj duży krok w tył, opuszczając biodra w dół, aż kolano nogi z tyłu znajdzie się tuż nad ziemią.", "czas_min": 3, "parametry": "3 serie x 10 powtórzeń na nogę", "miesnie": "Mięsień czworogłowy uda, mięsień pośladkowy wielki, m. dwugłowy"},
        {"nazwa": "Odwodzenie nogi z gumą", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Guma miniband nad kostkami. Stojąc na jednej nodze, odprowadzaj drugą nogę prostą kontrolowanym ruchem w bok, napinając pośladki.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń na stronę", "miesnie": "Mięsień pośladkowy średni, mięsień naprężacz powięzi szerokiej"},
        {"nazwa": "Przysiad bułgarski", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Jedna stopa oparta z tyłu o krzesło. Wykonaj przysiad na nogę wykrocznej, pilnując, by kolano nie uciekało do środka.", "czas_min": 3, "parametry": "3 serie x 8 powtórzeń na nogę", "miesnie": "Mięsień czworogłowy uda, mięsień pośladkowy wielki (praca unilateralna)"},
        {"nazwa": "Clamshell (Muszelka)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Leżenie na boku, kolana ugięte 90 stopni, stopy złączone. Unieś górne kolano w górę, nie rozłączając stóp ani miednicy.", "czas_min": 2, "parametry": "3 serie x 15 powtórzeń na stronę", "miesnie": "Mięsień pośladkowy średni, m. gruszkowaty, rotator zewnętrzny biodra"},
        {"nazwa": "Unoszenie prostej nogi (SLR)", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Jedna noga ugięta w kolanie. Drugą nogę (prostą w kolanie, palce zadarte na siebie) unoś do wysokości przeciwnego kolana.", "czas_min": 2, "parametry": "3 serie x 12 powtórzeń na nogę", "miesnie": "Mięsień prosty uda, mięsień napinacz powięzi szerokiej"},
        {"nazwa": "Martwy ciąg na jednej nodze", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Stojąc na jednej nodze, wykonaj skłon w przód z jednoczesnym wypchnięciem drugiej nogi w tył. Utrzymuj równe biodra.", "czas_min": 3, "parametry": "3 serie x 8 powtórzeń na stronę", "miesnie": "Mięsień pośladkowy wielki, grupa kulszowo-goleniowa, stabilizatory miednicy"},
        {"nazwa": "Krzesełko przy ścianie", "opis": "INSTRUKCJA RUCHU DLA PACJENTA: Oprzyj plecy o ścianę i zejdź do pozycji przysiadu (kąt 90 stopni w kolanach). Utrzymuj tę pozycję całkowicie statycznie.", "czas_min": 2, "parametry": "3 serie x 30 sekund", "miesnie": "Mięsień czworogłowy uda, mięsień pośladkowy wielki, m. przywodziciel wielki"}
    ]
}

# ==============================================================================
# 2. AUTORSKI ATLAS TRENINGU SPORTOWEGO NA SIŁOWNI
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

# ==============================================================================
# ŁADOWANIE WŁASNYCH ĆWICZEŃ Z PLIKU JSON
# ==============================================================================
PLIK_WLASNYCH_CWICZEN = "wlasne_cwiczenia.json"

def zaladuj_wlasne_cwiczenia():
    if os.path.exists(PLIK_WLASNYCH_CWICZEN):
        try:
            with open(PLIK_WLASNYCH_CWICZEN, "r", encoding="utf-8") as f:
                dane = json.load(f)
                for kat, lista in dane.get("FIZJO", {}).items():
                    if kat in BAZA_FIZJO:
                        BAZA_FIZJO[kat].extend(lista)
                for kat, lista in dane.get("GYM", {}).items():
                    if kat in BAZA_SILOWNIA:
                        BAZA_SILOWNIA[kat].extend(lista)
        except Exception as e:
            print(f"Błąd ładowania własnych ćwiczeń: {e}")

zaladuj_wlasne_cwiczenia()
GLOBALNA_BAZA = {**BAZA_FIZJO, **BAZA_SILOWNIA}

# ==============================================================================
# 3. INTERFEJS GRAFICZNY - CUSTOMTKINTER
# ==============================================================================
class FizjoWorkoutUltimateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Fizjo Workout Ultimate Pro - Modern Edition")
        self.root.geometry("1100x850")
        
        self.wylosowany_plan_cache = []  
        self.indeks_zaznaczonego_chwytaka = None
        self.ostatni_wybrany_profil = ""  
        self.realny_czas_cache = 0
        self.budzet_czasowy = 0
        self.is_manual_mode = False

        try:
            bg_image = ctk.CTkImage(Image.open("background.jpg"), size=(1920, 1080))
            self.bg_label = ctk.CTkLabel(self.root, text="", image=bg_image)
            self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        except Exception:
            pass

        self.main_frame = ctk.CTkFrame(self.root, fg_color=("#e0e0e0", "#141414"), corner_radius=15)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)

        self.btn_ai = ctk.CTkButton(self.main_frame, text="✨ Zapytaj AI", font=("Segoe UI", 12, "bold"),
                                    fg_color="#8e44ad", hover_color="#9b59b6", width=120, height=35,
                                    command=self.otworz_czat_ai)
        self.btn_ai.place(relx=0.98, rely=0.02, anchor="ne")
        
        lbl_t = ctk.CTkLabel(self.main_frame, text="FIZJO WORKOUT & GYM ENGINES", font=("Segoe UI", 26, "bold"), text_color="#2ecc71")
        lbl_t.pack(pady=(15, 2))
        
        lbl_s = ctk.CTkLabel(self.main_frame, text="Zintegrowane środowisko projektowania programów treningowych", font=("Segoe UI", 12), text_color="#a0a0a0")
        lbl_s.pack(pady=(0, 20))
        
        config_f = ctk.CTkFrame(self.main_frame, fg_color=("gray90", "#2b2b2b"), corner_radius=10)
        config_f.pack(fill=tk.X, padx=20, pady=(0, 15), ipadx=10, ipady=10)
        
        self.lbl_param_glowny = ctk.CTkLabel(config_f, text="Ilość ćw. NA PARTIĘ:", font=("Segoe UI", 12, "bold"))
        self.lbl_param_glowny.grid(row=0, column=0, padx=(10, 5), pady=10)
        
        self.entry_czas = ctk.CTkEntry(config_f, width=50, corner_radius=6, justify="center")
        self.entry_czas.insert(0, "4")  
        self.entry_czas.grid(row=0, column=1, padx=(0, 15))
        
        ctk.CTkLabel(config_f, text="Liczba Dni (Excel):", font=("Segoe UI", 12, "bold")).grid(row=0, column=2, padx=(5, 5))
        
        self.entry_dni = ctk.CTkEntry(config_f, width=50, corner_radius=6, justify="center")
        self.entry_dni.insert(0, "4")  
        self.entry_dni.grid(row=0, column=3, padx=(0, 15))
        
        ctk.CTkLabel(config_f, text="Profil Silnika:", font=("Segoe UI", 12, "bold")).grid(row=0, column=4, padx=(5, 5))
        
        self.combo_tryb = ctk.CTkComboBox(config_f, width=350, corner_radius=6, values=[
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
        ], command=self.odswiez_labele_konfiguracji)
        self.combo_tryb.set("GYM: Automatyczny Split (Dni Tygodnia)")
        self.combo_tryb.grid(row=0, column=5, padx=5)
        
        self.lbl_licznik_czasu = ctk.CTkLabel(config_f, text="Gotowy do generowania planu...", font=("Segoe UI", 12, "bold"), text_color="#a0a0a0")
        self.lbl_licznik_czasu.grid(row=1, column=0, columnspan=6, sticky="w", padx=10, pady=(5, 5))
        
        buttons_f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        buttons_f.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        self.btn_automat = ctk.CTkButton(buttons_f, text="⚡ AUTOMAT", font=("Segoe UI", 12, "bold"),
                                         fg_color="#2ecc71", text_color="#0f0f0f", hover_color="#27ae60",
                                         corner_radius=8, command=self.generuj_plan_automatycznie)
        self.btn_automat.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5), ipady=5)
        
        self.btn_manual = ctk.CTkButton(buttons_f, text="＋ KREATOR", font=("Segoe UI", 12, "bold"),
                                        fg_color="#16a085", hover_color="#117a65",
                                        corner_radius=8, command=self.otworz_kreator_wyboru)
        self.btn_manual.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, ipady=5)

        self.btn_edytuj_baza = ctk.CTkButton(buttons_f, text="⚙️ EDYTUJ BAZĘ", font=("Segoe UI", 12, "bold"),
                                             fg_color="#d35400", hover_color="#e67e22",
                                             corner_radius=8, command=self.otworz_menedzer_bazy)
        self.btn_edytuj_baza.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, ipady=5)
        
        self.btn_zapisz = ctk.CTkButton(buttons_f, text="💾 ZAPISZ DOC", font=("Segoe UI", 12, "bold"),
                                        fg_color="#2980b9", hover_color="#3498db", state="disabled",
                                        corner_radius=8, command=self.zapisz_do_docx)
        self.btn_zapisz.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, ipady=5)
        
        self.btn_zapisz_excel = ctk.CTkButton(buttons_f, text="📊 ZAPISZ EXCEL", font=("Segoe UI", 12, "bold"),
                                              fg_color="#8e44ad", hover_color="#9b59b6", state="disabled",
                                              corner_radius=8, command=self.zapisz_do_excel)
        self.btn_zapisz_excel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, ipady=5)

        self.btn_reset = ctk.CTkButton(buttons_f, text="Reset X", font=("Segoe UI", 12, "bold"), width=80,
                                       fg_color="#c0392b", hover_color="#e74c3c",
                                       corner_radius=8, command=self.odswiez_widok_ekranu)
        self.btn_reset.pack(side=tk.LEFT, padx=(5, 0), ipady=5)

        self.scrollable_f = ctk.CTkScrollableFrame(self.main_frame, corner_radius=10, fg_color=("#d0d0d0", "#242424"))
        self.scrollable_f.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        self.odswiez_labele_konfiguracji()
        self.odswiez_widok_ekranu()

    def odswiez_labele_konfiguracji(self, event=None):
        profil = self.combo_tryb.get()
        if profil.startswith("GYM:"):
            self.lbl_param_glowny.configure(text="Ilość ćw. NA PARTIĘ:")
            try:
                val = int(self.entry_czas.get() or "0")
                if val >= 15:
                    self.entry_czas.delete(0, tk.END)
                    self.entry_czas.insert(0, "4")
            except ValueError:
                pass
        else:
            self.lbl_param_glowny.configure(text="Czas sesji (min):")
            try:
                val = int(self.entry_czas.get() or "0")
                if val <= 10:
                    self.entry_czas.delete(0, tk.END)
                    self.entry_czas.insert(0, "45")
            except ValueError:
                pass

    def czysc_ekran_skrolla(self):
        for widget in self.scrollable_f.winfo_children():
            widget.destroy()

    def odswiez_widok_ekranu(self):
        self.wylosowany_plan_cache.clear()
        self.realny_czas_cache = 0
        self.budzet_czasowy = 0
        self.is_manual_mode = False
        self.indeks_zaznaczonego_chwytaka = None
        self.ostatni_wybrany_profil = ""
        self.czysc_ekran_skrolla()
        
        self.btn_zapisz.configure(state="disabled", fg_color="#2c3e50")
        self.btn_zapisz_excel.configure(state="disabled", fg_color="#2c3e50")
        self.btn_automat.configure(state="normal", fg_color="#2ecc71")
        self.btn_manual.configure(state="normal", fg_color="#16a085")
        self.entry_czas.configure(state="normal")
        self.entry_dni.configure(state="normal")
        self.combo_tryb.configure(state="readonly")
        
        self.lbl_licznik_czasu.configure(text="Gotowy do generowania planu...", text_color="#7f8c8d")
        
        welcome_lbl = ctk.CTkLabel(
            self.scrollable_f, 
            text="Ustal parametry sesji oraz liczbę dni do Excela na górnym pasku.\nWybierz profil z menu rozwijanego i uruchom AUTOMAT, lub stwórz plan manualnie.", 
            font=("Segoe UI", 14), text_color="#7f8c8d"
        )
        welcome_lbl.pack(fill=tk.X, expand=True, pady=50)

    # ==============================================================================
    # MENEDŻER BAZY
    # ==============================================================================
    def otworz_okno_dodawania_cwiczenia(self, edit_data=None):
        okno = ctk.CTkToplevel(self.root)
        tytul = "Edytuj własne ćwiczenie" if edit_data else "Dodaj własne ćwiczenie do bazy"
        okno.title(tytul)
        okno.geometry("500x600")
        okno.configure(fg_color="#141414")
        okno.attributes('-topmost', True) 
        
        ctk.CTkLabel(okno, text=tytul.upper(), font=("Segoe UI", 18, "bold"), text_color="#8e44ad").pack(pady=15)
        
        form_frame = ctk.CTkFrame(okno, fg_color="transparent")
        form_frame.pack(fill=tk.BOTH, expand=True, padx=20)
        
        def create_label(text):
            lbl = ctk.CTkLabel(form_frame, text=text, font=("Segoe UI", 12, "bold"), text_color="#eceff1")
            lbl.pack(anchor="w", pady=(10, 2))
            return lbl
            
        def create_entry():
            e = ctk.CTkEntry(form_frame, font=("Segoe UI", 13), width=400, corner_radius=6)
            e.pack(fill=tk.X)
            return e
            
        create_label("Kategoria Docelowa:")
        kategorie = [f"FIZJO: {k}" for k in BAZA_FIZJO.keys()] + [f"GYM: {k}" for k in BAZA_SILOWNIA.keys()]
        combo_kat = ctk.CTkComboBox(form_frame, font=("Segoe UI", 13), values=kategorie, corner_radius=6)
        combo_kat.pack(fill=tk.X)
        
        create_label("Nazwa ćwiczenia:")
        e_nazwa = create_entry()
        
        create_label("Zalecenie (np. 3x12, 5 minut):")
        e_parametry = create_entry()
        
        create_label("Szacowany czas (w minutach, cyfra):")
        e_czas = create_entry()
        
        create_label("Anatomia (pracujące mięśnie):")
        e_miesnie = create_entry()
        
        create_label("Opis / Instrukcja ruchu:")
        t_opis = ctk.CTkTextbox(form_frame, height=80, font=("Segoe UI", 13), corner_radius=6)
        t_opis.pack(fill=tk.X)

        if edit_data:
            typ_bazy, kategoria, idx, cw_dane = edit_data
            combo_kat.set(f"{typ_bazy}: {kategoria}")
            combo_kat.configure(state="disabled") 
            
            e_nazwa.insert(0, cw_dane['nazwa'])
            e_parametry.insert(0, cw_dane['parametry'])
            e_czas.insert(0, str(cw_dane['czas_min']))
            e_miesnie.insert(0, cw_dane['miesnie'])
            t_opis.insert("1.0", cw_dane['opis'])
        else:
            if kategorie: combo_kat.set(kategorie[0])
            e_czas.insert(0, "2")
        
        def zapisz_cwiczenie():
            kat_pelna = combo_kat.get()
            nazwa = e_nazwa.get().strip()
            parametry = e_parametry.get().strip()
            czas = e_czas.get().strip()
            miesnie = e_miesnie.get().strip()
            opis = t_opis.get("1.0", tk.END).strip()
            
            if not nazwa or not parametry or not czas or not miesnie or not opis:
                messagebox.showwarning("Błąd", "Wypełnij wszystkie pola formularza.")
                return
                
            try:
                czas_int = int(czas)
            except ValueError:
                messagebox.showwarning("Błąd", "Czas musi być podany jako liczba całkowita (np. 2).")
                return
                
            nowe_cw = {
                "nazwa": nazwa,
                "opis": opis,
                "czas_min": czas_int,
                "parametry": parametry,
                "miesnie": miesnie
            }
            
            typ_bazy_str, kategoria_str = kat_pelna.split(": ")
            dane_z_pliku = {"FIZJO": {}, "GYM": {}}
            
            if os.path.exists(PLIK_WLASNYCH_CWICZEN):
                try:
                    with open(PLIK_WLASNYCH_CWICZEN, "r", encoding="utf-8") as f:
                        dane_z_pliku = json.load(f)
                except:
                    pass
            
            if edit_data:
                stara_nazwa = cw_dane['nazwa']
                dane_z_pliku[typ_bazy_str][kategoria_str][idx] = nowe_cw
                baza_docelowa = BAZA_FIZJO if typ_bazy_str == "FIZJO" else BAZA_SILOWNIA
                if kategoria_str in baza_docelowa:
                    for i_mem, c in enumerate(baza_docelowa[kategoria_str]):
                        if c['nazwa'] == stara_nazwa:
                            baza_docelowa[kategoria_str][i_mem] = nowe_cw
                            break
                            
                if kategoria_str in GLOBALNA_BAZA:
                    for i_mem, c in enumerate(GLOBALNA_BAZA[kategoria_str]):
                        if c['nazwa'] == stara_nazwa:
                            GLOBALNA_BAZA[kategoria_str][i_mem] = nowe_cw
                            break
                msg = "Zmiany zostały poprawnie zapisane!"
            else:
                if typ_bazy_str not in dane_z_pliku:
                    dane_z_pliku[typ_bazy_str] = {}
                if kategoria_str not in dane_z_pliku[typ_bazy_str]:
                    dane_z_pliku[typ_bazy_str][kategoria_str] = []
                    
                dane_z_pliku[typ_bazy_str][kategoria_str].append(nowe_cw)
                
                if typ_bazy_str == "FIZJO":
                    BAZA_FIZJO[kategoria_str].append(nowe_cw)
                else:
                    BAZA_SILOWNIA[kategoria_str].append(nowe_cw)
                GLOBALNA_BAZA[kategoria_str].append(nowe_cw)
                msg = "Ćwiczenie zostało trwale dodane do Twojej bazy!"
            
            try:
                with open(PLIK_WLASNYCH_CWICZEN, "w", encoding="utf-8") as f:
                    json.dump(dane_z_pliku, f, ensure_ascii=False, indent=4)
                messagebox.showinfo("Sukces", msg)
                okno.destroy()
                if hasattr(self, 'okno_menedzera') and self.okno_menedzera.winfo_exists():
                    self.odswiez_widok_menedzera()
            except Exception as e:
                messagebox.showerror("Błąd", f"Nie udało się zapisać do pliku: {e}")

        btn_zapisz = ctk.CTkButton(okno, text="ZAPISZ DO BAZY", font=("Segoe UI", 14, "bold"), fg_color="#8e44ad", hover_color="#9b59b6", height=40, command=zapisz_cwiczenie)
        btn_zapisz.pack(fill=tk.X, padx=20, pady=20)

    def otworz_menedzer_bazy(self):
        if not os.path.exists(PLIK_WLASNYCH_CWICZEN):
            messagebox.showinfo("Baza", "Nie masz jeszcze dodanych własnych ćwiczeń.")
            return

        with open(PLIK_WLASNYCH_CWICZEN, "r", encoding="utf-8") as f:
            try:
                dane = json.load(f)
            except:
                dane = {}

        if not dane.get("FIZJO") and not dane.get("GYM"):
            messagebox.showinfo("Baza", "Baza własnych ćwiczeń jest pusta.")
            return

        self.okno_menedzera = ctk.CTkToplevel(self.root)
        self.okno_menedzera.title("Menedżer własnych ćwiczeń")
        self.okno_menedzera.geometry("750x600")
        self.okno_menedzera.configure(fg_color="#141414")
        self.okno_menedzera.attributes('-topmost', True)
        
        ctk.CTkLabel(self.okno_menedzera, text="ZARZĄDZANIE WŁASNYMI ĆWICZENIAMI", font=("Segoe UI", 18, "bold"), text_color="#e67e22").pack(pady=15)

        self.menedzer_scroll = ctk.CTkScrollableFrame(self.okno_menedzera, corner_radius=10, fg_color=("#d0d0d0", "#1e1e1e"))
        self.menedzer_scroll.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        self.odswiez_widok_menedzera()

    def odswiez_widok_menedzera(self):
        for widget in self.menedzer_scroll.winfo_children():
            widget.destroy()

        if not os.path.exists(PLIK_WLASNYCH_CWICZEN):
            return

        with open(PLIK_WLASNYCH_CWICZEN, "r", encoding="utf-8") as f:
            try:
                dane = json.load(f)
            except:
                dane = {}

        pusto = True
        for typ_bazy in ["FIZJO", "GYM"]:
            if typ_bazy in dane:
                for kat, lista_cw in dane[typ_bazy].items():
                    if lista_cw:
                        pusto = False
                        ctk.CTkLabel(self.menedzer_scroll, text=f"■ {typ_bazy} - {kat.upper()}", font=("Segoe UI", 14, "bold"), text_color="#7f8c8d").pack(anchor="w", pady=(15, 5))
                        
                        for idx, cw in enumerate(lista_cw):
                            c_frame = ctk.CTkFrame(self.menedzer_scroll, fg_color="#2b2b2b", corner_radius=6)
                            c_frame.pack(fill=tk.X, pady=3, padx=5)
                            
                            ctk.CTkLabel(c_frame, text=f"{cw['nazwa']}", font=("Segoe UI", 14, "bold"), text_color="#ffffff").pack(side=tk.LEFT, padx=10, pady=8)
                            ctk.CTkLabel(c_frame, text=f"({cw['czas_min']} min)", font=("Segoe UI", 12), text_color="#bdc3c7").pack(side=tk.LEFT)
                            
                            btn_usun = ctk.CTkButton(c_frame, text="Usuń X", font=("Segoe UI", 11, "bold"), fg_color="#c0392b", hover_color="#e74c3c", width=60, height=26,
                                command=lambda tb=typ_bazy, k=kat, i=idx, n=cw['nazwa']: self.usun_wlasne_cw_z_bazy(tb, k, i, n))
                            btn_usun.pack(side=tk.RIGHT, padx=(5, 10))
                            
                            btn_edycja = ctk.CTkButton(c_frame, text="Edytuj ✎", font=("Segoe UI", 11, "bold"), fg_color="#f1c40f", text_color="#000000", hover_color="#f39c12", width=60, height=26,
                                command=lambda tb=typ_bazy, k=kat, i=idx, data=cw: self.otworz_okno_dodawania_cwiczenia((tb, k, i, data)))
                            btn_edycja.pack(side=tk.RIGHT)

        if pusto:
            ctk.CTkLabel(self.menedzer_scroll, text="Brak dodanych ćwiczeń.", text_color="#7f8c8d").pack(pady=20)

    def usun_wlasne_cw_z_bazy(self, typ_bazy, kat, idx, nazwa):
        if messagebox.askyesno("Potwierdzenie", f"Czy na pewno chcesz usunąć to ćwiczenie:\n\n'{nazwa}'\n\nz bazy danych?"):
            with open(PLIK_WLASNYCH_CWICZEN, "r", encoding="utf-8") as f:
                dane = json.load(f)
            
            del dane[typ_bazy][kat][idx]
            if not dane[typ_bazy][kat]: del dane[typ_bazy][kat]
                
            with open(PLIK_WLASNYCH_CWICZEN, "w", encoding="utf-8") as f:
                json.dump(dane, f, ensure_ascii=False, indent=4)
                
            baza_docelowa = BAZA_FIZJO if typ_bazy == "FIZJO" else BAZA_SILOWNIA
            if kat in baza_docelowa:
                baza_docelowa[kat] = [c for c in baza_docelowa[kat] if c['nazwa'] != nazwa]
            if kat in GLOBALNA_BAZA:
                GLOBALNA_BAZA[kat] = [c for c in GLOBALNA_BAZA[kat] if c['nazwa'] != nazwa]
                
            self.odswiez_widok_menedzera()

    # ==============================================================================
    # GENERATOR I LOGIKA KAFELKOWA
    # ==============================================================================
    def generuj_plan_automatycznie(self):
        try:
            self.budzet_czasowy = int(self.entry_czas.get())
            if self.budzet_czasowy <= 0: raise ValueError
        except ValueError:
            messagebox.showerror("Błąd", "Wprowadź prawidłową wartość w pierwszym polu.")
            return

        self.czysc_ekran_skrolla()
        self.is_manual_mode = False
        self.wylosowany_plan_cache = []
        self.realny_czas_cache = 0
        
        profil = self.combo_tryb.get()
        self.ostatni_wybrany_profil = profil
        
        b_fizjo = {k: list(v) for k, v in BAZA_FIZJO.items()}
        b_gym = {k: list(v) for k, v in BAZA_SILOWNIA.items()}

        if profil.startswith("FIZJO:"):
            if "Kompleksowy" in profil:
                oddech_start_org = random.choice(b_fizjo["Oddechowe"])
                oddech_start = oddech_start_org.copy()
                oddech_start["uwagi"] = ""
                b_fizjo["Oddechowe"].remove(oddech_start_org)
                
                self.wylosowany_plan_cache.append(("Oddechowe (Rozgrzewka)", oddech_start))
                self.realny_czas_cache += oddech_start["czas_min"]
                
                oddech_koniec_org = random.choice(b_fizjo["Oddechowe"]) if b_fizjo["Oddechowe"] else None
                if oddech_koniec_org:
                    oddech_koniec = oddech_koniec_org.copy()
                    oddech_koniec["uwagi"] = ""
                    czas_koncowy = oddech_koniec["czas_min"]
                else:
                    oddech_koniec = None
                    czas_koncowy = 0

                sztywny_lancuch = ["Głowa/Szyja", "Kończyna górna", "Core (Tułów)", "Kończyna dolna"]
                max_rund = 10
                puste = 0
                
                while self.realny_czas_cache + czas_koncowy < self.budzet_czasowy and puste < max_rund:
                    dodano = False
                    for kat in sztywny_lancuch:
                        if len(b_fizjo[kat]) > 0:
                            cw_org = random.choice(b_fizjo[kat])
                            if self.realny_czas_cache + cw_org["czas_min"] + czas_koncowy <= self.budzet_czasowy:
                                cw = cw_org.copy()
                                cw["uwagi"] = ""
                                self.wylosowany_plan_cache.append((kat, cw))
                                self.realny_czas_cache += cw["czas_min"]
                                b_fizjo[kat].remove(cw_org)
                                dodano = True
                    if not dodano: puste += 1
                    else: puste = 0

                if oddech_koniec:
                    self.wylosowany_plan_cache.append(("Oddechowe (Wyciszenie)", oddech_koniec))
                    self.realny_czas_cache += oddech_koniec["czas_min"]

                waga = {
                    "Oddechowe (Rozgrzewka)": 0, "Głowa/Szyja": 1, "Kończyna górna": 2, 
                    "Core (Tułów)": 3, "Kończyna dolna": 4, "Oddechowe (Wyciszenie)": 5
                }
                self.wylosowany_plan_cache.sort(key=lambda x: waga.get(x[0], 99))

            else:
                wybrana_partia_fizjo = profil.split(" - ")[1].replace("Tylko ", "")
                dostepne_cwiczenia = b_fizjo[wybrana_partia_fizjo]
                
                while self.realny_czas_cache < self.budzet_czasowy and dostepne_cwiczenia:
                    cw_org = random.choice(dostepne_cwiczenia)
                    if self.realny_czas_cache + cw_org["czas_min"] <= self.budzet_czasowy:
                        cw = cw_org.copy()
                        cw["uwagi"] = ""
                        self.wylosowany_plan_cache.append((wybrana_partia_fizjo, cw))
                        self.realny_czas_cache += cw["czas_min"]
                        dostepne_cwiczenia.remove(cw_org)
                    else:
                        break

        elif profil.startswith("GYM:"):
            # W TYM TRYBIE BUDŻET TO ŚCISŁA ILOŚĆ ĆWICZEŃ NA DANĄ PARTIĘ
            limit_cw_na_partie = self.budzet_czasowy
            
            # Pobieramy rozgrzewkę i zakończenie dla każdego profilu siłowego
            rozgrzewka_gym_org = random.choice(b_gym["Rozgrzewka"])
            rozgrzewka_gym = rozgrzewka_gym_org.copy()
            rozgrzewka_gym["uwagi"] = ""
            
            zakonczenie_gym_org = random.choice(b_gym["Zakończenie treningu"])
            zakonczenie_gym = zakonczenie_gym_org.copy()
            zakonczenie_gym["uwagi"] = ""

            if profil == "GYM: Automatyczny Split (Dni Tygodnia)":
                try:
                    liczba_dni = int(self.entry_dni.get())
                except ValueError:
                    liczba_dni = 3
                
                if liczba_dni < 1: liczba_dni = 1
                if liczba_dni > 7: liczba_dni = 7
                
                dni_tygodnia = ["Poniedziałek", "Wtorek", "Środa", "Czwartek", "Piątek", "Sobota", "Niedziela"]
                
                if liczba_dni == 1: plan_na_dni = [["Klatka piersiowa", "Plecy", "Nogi", "Ręce", "Pośladki"]]
                elif liczba_dni == 2: plan_na_dni = [["Klatka piersiowa", "Ręce"], ["Plecy", "Nogi", "Pośladki"]]
                elif liczba_dni == 3: plan_na_dni = [["Klatka piersiowa", "Ręce"], ["Nogi", "Pośladki"], ["Plecy"]]
                elif liczba_dni == 4: plan_na_dni = [["Ręce"], ["Klatka piersiowa"], ["Nogi"], ["Plecy"]]
                elif liczba_dni == 5: plan_na_dni = [["Klatka piersiowa"], ["Plecy"], ["Nogi"], ["Ręce"], ["Pośladki"]]
                elif liczba_dni == 6: plan_na_dni = [["Klatka piersiowa"], ["Plecy"], ["Nogi"], ["Ręce"], ["Pośladki"], ["Klatka piersiowa", "Ręce"]]
                else: plan_na_dni = [["Klatka piersiowa"], ["Plecy"], ["Nogi"], ["Ręce"], ["Pośladki"], ["Klatka piersiowa"], ["Plecy"]]

                for i in range(liczba_dni):
                    dzien_nazwa = dni_tygodnia[i]
                    partie_w_dniu = plan_na_dni[i]
                    nazwa_w_naglowku = " + ".join([p.upper() for p in partie_w_dniu])
                    
                    cw_naglowek = {
                        "nazwa": f"{dzien_nazwa}: {nazwa_w_naglowku}", 
                        "opis": f"Trening ukierunkowany na: {', '.join(partie_w_dniu)}.", 
                        "czas_min": 0, "parametry": "-", "miesnie": "-", "uwagi": ""
                    }
                    self.wylosowany_plan_cache.append(("NAGŁÓWEK DNIA", cw_naglowek))
                    self.wylosowany_plan_cache.append(("GYM: Rozgrzewka", rozgrzewka_gym.copy()))
                    
                    for partia in partie_w_dniu:
                        dostepne = list(b_gym[partia])
                        dodano_dla_partii = 0
                        
                        while dodano_dla_partii < limit_cw_na_partie and dostepne:
                            cw_org = random.choice(dostepne)
                            cw_copy = cw_org.copy()
                            cw_copy["uwagi"] = "" 
                            self.wylosowany_plan_cache.append((f"GYM: {partia}", cw_copy))
                            dodano_dla_partii += 1
                            dostepne.remove(cw_org)
                            
                    self.wylosowany_plan_cache.append(("GYM: Zakończenie", zakonczenie_gym.copy()))

            else:
                self.wylosowany_plan_cache.append(("GYM: Rozgrzewka", rozgrzewka_gym))

                if "Ogólnorozwojowy" in profil:
                    kolejnosc_fbw = ["Klatka piersiowa", "Plecy", "Nogi", "Ręce", "Pośladki"]
                    
                    for partia in kolejnosc_fbw:
                        dostepne = list(b_gym[partia])
                        dodano = 0
                        while dodano < limit_cw_na_partie and dostepne:
                            cw_org = random.choice(dostepne)
                            cw = cw_org.copy()
                            cw["uwagi"] = ""
                            self.wylosowany_plan_cache.append((f"GYM: {partia}", cw))
                            b_gym[partia].remove(cw_org) 
                            dodano += 1
                else:
                    wybrana_partia = profil.split(" - ")[1]
                    dostepne_cwiczenia = b_gym[wybrana_partia]
                    dodano = 0
                    
                    while dodano < limit_cw_na_partie and dostepne_cwiczenia:
                        cw_org = random.choice(dostepne_cwiczenia)
                        cw = cw_org.copy()
                        cw["uwagi"] = ""
                        self.wylosowany_plan_cache.append((f"GYM: {wybrana_partia}", cw))
                        dodano += 1
                        dostepne_cwiczenia.remove(cw_org)

                self.wylosowany_plan_cache.append(("GYM: Zakończenie", zakonczenie_gym))

        self.entry_czas.configure(state="disabled")
        self.entry_dni.configure(state="disabled")
        self.combo_tryb.configure(state="disabled")
        self.renderuj_kafelki_ekranu()

    def otworz_kreator_wyboru(self):
        try:
            if self.realny_czas_cache == 0:  
                self.budzet_czasowy = int(self.entry_czas.get())
                if self.budzet_czasowy <= 0: raise ValueError
                self.entry_czas.configure(state="disabled")
                self.entry_dni.configure(state="disabled")
                self.combo_tryb.configure(state="disabled")
                self.btn_automat.configure(state="disabled", fg_color="#2c3e50")
        except ValueError:
            messagebox.showerror("Błąd", "Wprowadź prawidłową wartość bazową w polu.")
            return

        self.is_manual_mode = True
        self.ostatni_wybrany_profil = self.combo_tryb.get()
        
        okno_wyboru = ctk.CTkToplevel(self.root)
        okno_wyboru.title("Kreator swobodny (Bez limitu)")
        okno_wyboru.geometry("750x800")
        okno_wyboru.configure(fg_color="#141414")
        okno_wyboru.attributes('-topmost', True)

        ctk.CTkLabel(okno_wyboru, text="INTEGROWANY ATLAS MEDYCZNO-TRENINGOWY", font=("Segoe UI", 18, "bold"), text_color="#2ecc71").pack(pady=10)
        ctk.CTkLabel(okno_wyboru, text="Tryb kreatora aktywny - przydzielaj ćwiczenia swobodnie do planu", font=("Segoe UI", 12, "bold"), text_color="#ffffff").pack(pady=(0, 10))

        filter_p = ctk.CTkFrame(okno_wyboru, corner_radius=10, fg_color="#2b2b2b")
        filter_p.pack(fill=tk.X, padx=20, pady=(0, 10), ipady=5)

        lista_zakladek = [
            "Fizjo: Oddechowe", "Fizjo: Głowa/Szyja", "Fizjo: Kończyna górna", 
            "Fizjo: Core (Tułów)", "Fizjo: Kończyna dolna", "GYM: Rozgrzewka", 
            "GYM: Klatka piersiowa", "GYM: Plecy", "GYM: Ręce", "GYM: Nogi", 
            "GYM: Pośladki", "GYM: Zakończenie"
        ]
        
        ctk.CTkLabel(filter_p, text="Filtruj kategorię:", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, padx=(10,5), pady=10)
        combo_kat = ctk.CTkComboBox(filter_p, values=lista_zakladek, width=200)
        combo_kat.grid(row=0, column=1, padx=5)
        combo_kat.set(lista_zakladek[0])

        ctk.CTkLabel(filter_p, text="Szukaj mięśnia:", font=("Segoe UI", 12, "bold")).grid(row=0, column=2, padx=(15, 5))
        entry_szukaj_miesnia = ctk.CTkEntry(filter_p, width=150)
        entry_szukaj_miesnia.grid(row=0, column=3, padx=5)

        self.kreator_scroll = ctk.CTkScrollableFrame(okno_wyboru, corner_radius=10, fg_color=("#d0d0d0", "#1e1e1e"))
        self.kreator_scroll.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        wybrane_w_kreatorze_nazwy = {x[1]['nazwa'] for x in self.wylosowany_plan_cache}
        zmienne_checkboxow = {}

        def odswiez_liste_wyboru(*args):
            for nazwa, (var, kat, cw) in list(zmienne_checkboxow.items()):
                if var.get(): wybrane_w_kreatorze_nazwy.add(nazwa)
                else: wybrane_w_kreatorze_nazwy.discard(nazwa)

            for w in self.kreator_scroll.winfo_children(): w.destroy()
            zmienne_checkboxow.clear()
            
            fraza = entry_szukaj_miesnia.get().strip().lower()
            zakladka = combo_kat.get()

            def uprosc(t): 
                return t.lower().replace("ó","o").replace("ł","l").replace("ś","s")\
                        .replace("ą","a").replace("ę","e").replace("ć","c")\
                        .replace("ź","z").replace("ż","z").replace("ń","n")
            f_czysta = uprosc(fraza)

            pula = []
            if f_czysta:
                for k, v in BAZA_FIZJO.items():
                    for cw in v:
                        if f_czysta in uprosc(cw["miesnie"]): pula.append((k, cw))
                for k, v in BAZA_SILOWNIA.items():
                    for cw in v:
                        if f_czysta in uprosc(cw["miesnie"]): pula.append((f"GYM: {k}", cw))
            else:
                if zakladka.startswith("Fizjo: "):
                    kat_czysta = zakladka.replace("Fizjo: ", "")
                    for cw in BAZA_FIZJO[kat_czysta]: pula.append((kat_czysta, cw))
                else:
                    kat_czysta = zakladka.replace("GYM: ", "").replace("Zakończenie", "Zakończenie treningu")
                    for cw in BAZA_SILOWNIA[kat_czysta]: pula.append((f"GYM: {kat_czysta}", cw))

            for kat, cw in pula:
                var = tk.BooleanVar(value=cw['nazwa'] in wybrane_w_kreatorze_nazwy)
                zmienne_checkboxow[cw['nazwa']] = (var, kat, cw)
                
                chk_tekst = f"{cw['nazwa']} ({cw['czas_min']} min)\n↳ [{cw['miesnie']}]"
                chk = ctk.CTkCheckBox(self.kreator_scroll, text=chk_tekst, variable=var, font=("Segoe UI", 12), text_color="#ffffff", fg_color="#2ecc71")
                chk.pack(anchor="w", padx=10, pady=8)

        combo_kat.configure(command=odswiez_liste_wyboru)
        entry_szukaj_miesnia.bind("<KeyRelease>", odswiez_liste_wyboru)
        odswiez_liste_wyboru()

        def zatwierdz_wybor():
            for nazwa, (var, kat, cw) in zmienne_checkboxow.items():
                if var.get(): wybrane_w_kreatorze_nazwy.add(nazwa)
                else: wybrane_w_kreatorze_nazwy.discard(nazwa)

            nowy_plan = []
            
            for kat, cwiczenia in {**BAZA_FIZJO, **BAZA_SILOWNIA}.items():
                for cw in cwiczenia:
                    if cw['nazwa'] in wybrane_w_kreatorze_nazwy:
                        etykieta = f"GYM: {kat}" if kat in BAZA_SILOWNIA else kat
                        cw_kopiowane = cw.copy()
                        cw_kopiowane["uwagi"] = "" 
                        nowy_plan.append((etykieta, cw_kopiowane))
            
            self.wylosowany_plan_cache = nowy_plan
            self.renderuj_kafelki_ekranu()
            okno_wyboru.destroy()

        btn_zatwierdz = ctk.CTkButton(okno_wyboru, text="ZATWIERDŹ PROGRAM (BEZ LIMITÓW)", font=("Segoe UI", 14, "bold"), fg_color="#2ecc71", text_color="#0f0f0f", hover_color="#27ae60", height=45, command=zatwierdz_wybor)
        btn_zatwierdz.pack(fill=tk.X, padx=20, pady=20)

    def obsluga_klikniecia_chwytaka(self, index):
        if self.indeks_zaznaczonego_chwytaka is None:
            self.indeks_zaznaczonego_chwytaka = index
            self.renderuj_kafelki_ekranu()
        elif self.indeks_zaznaczonego_chwytaka == index:
            self.indeks_zaznaczonego_chwytaka = None
            self.renderuj_kafelki_ekranu()
        else:
            idx1, idx2 = self.indeks_zaznaczonego_chwytaka, index
            self.wylosowany_plan_cache[idx1], self.wylosowany_plan_cache[idx2] = \
                self.wylosowany_plan_cache[idx2], self.wylosowany_plan_cache[idx1]
            self.indeks_zaznaczonego_chwytaka = None
            self.renderuj_kafelki_ekranu()

    def usun_konkretne_cwiczenie(self, index):
        self.wylosowany_plan_cache.pop(index)
        self.indeks_zaznaczonego_chwytaka = None
        self.renderuj_kafelki_ekranu()

    def edytuj_parametry_w_planie(self, index):
        kat, cw = self.wylosowany_plan_cache[index]
        if kat == "NAGŁÓWEK DNIA":
            return
            
        okno_edycji = ctk.CTkToplevel(self.root)
        okno_edycji.title("Edycja parametrów i uwag")
        okno_edycji.geometry("450x320")
        okno_edycji.configure(fg_color="#141414")
        okno_edycji.attributes('-topmost', True)

        ctk.CTkLabel(okno_edycji, text=f"Edytuj: {cw['nazwa']}", font=("Segoe UI", 16, "bold"), text_color="#2ecc71", wraplength=400).pack(pady=15)

        ctk.CTkLabel(okno_edycji, text="Parametry (np. 4x10, 3 serie, 15 min):", font=("Segoe UI", 12)).pack(pady=(5, 0))
        entry_param = ctk.CTkEntry(okno_edycji, width=350, font=("Segoe UI", 13))
        entry_param.insert(0, cw.get('parametry', ''))
        entry_param.pack(pady=5)
        
        ctk.CTkLabel(okno_edycji, text="Twoje uwagi do ćwiczenia (opcjonalnie):", font=("Segoe UI", 12)).pack(pady=(15, 0))
        entry_uwagi = ctk.CTkEntry(okno_edycji, width=350, font=("Segoe UI", 13))
        entry_uwagi.insert(0, cw.get('uwagi', ''))
        entry_uwagi.pack(pady=5)

        def zapisz():
            nowe_parametry = entry_param.get().strip()
            nowe_uwagi = entry_uwagi.get().strip()
            if nowe_parametry:
                self.wylosowany_plan_cache[index][1]['parametry'] = nowe_parametry
                self.wylosowany_plan_cache[index][1]['uwagi'] = nowe_uwagi
                self.renderuj_kafelki_ekranu()
                okno_edycji.destroy()
            else:
                messagebox.showwarning("Błąd", "Parametry nie mogą być puste.")

        btn_zapisz = ctk.CTkButton(okno_edycji, text="ZAPISZ", font=("Segoe UI", 14, "bold"), fg_color="#2ecc71", text_color="#0f0f0f", hover_color="#27ae60", height=40, command=zapisz)
        btn_zapisz.pack(pady=20, fill=tk.X, padx=50)

    def renderuj_kafelki_ekranu(self):
        self.czysc_ekran_skrolla()
        is_gym = self.ostatni_wybrany_profil.startswith("GYM:")
        naglowek_info = "ZESTAW AUTOMATYCZNY" if not self.is_manual_mode else "MANUALNY PROTOKÓŁ KREATORA"

        if is_gym:
            ilosc_glownych = sum(1 for k, cw in self.wylosowany_plan_cache if k not in ["NAGŁÓWEK DNIA", "GYM: Rozgrzewka", "GYM: Zakończenie"])
            self.lbl_licznik_czasu.configure(
                text=f"{naglowek_info} | Wygenerowano {ilosc_glownych} głównych ćwiczeń.", 
                text_color="#2ecc71"
            )
        else:
            self.realny_czas_cache = sum(x[1].get('czas_min', 0) for x in self.wylosowany_plan_cache)
            minuty_zostaly = self.budzet_czasowy - self.realny_czas_cache
            self.lbl_licznik_czasu.configure(
                text=f"{naglowek_info} | Rzeczywisty czas: {self.realny_czas_cache} / {self.budzet_czasowy} min (Zostało: {minuty_zostaly} min)", 
                text_color="#2ecc71" if minuty_zostaly >= 0 else "#e74c3c"
            )
        
        if self.wylosowany_plan_cache: 
            self.btn_zapisz.configure(state="normal", fg_color="#2980b9")
            self.btn_zapisz_excel.configure(state="normal", fg_color="#8e44ad")
        else: 
            self.odswiez_widok_ekranu()
            return

        for idx, (kategoria, cw) in enumerate(self.wylosowany_plan_cache):
            czy_zaznaczony = (self.indeks_zaznaczonego_chwytaka == idx)
            
            tile_bg = "#2c3e50" if czy_zaznaczony else "#2b2b2b"
            border_color = "#2ecc71" if czy_zaznaczony else "#3a3a3a"

            tile = ctk.CTkFrame(self.scrollable_f, fg_color=tile_bg, border_width=2, border_color=border_color, corner_radius=10)
            tile.pack(fill=tk.X, pady=8, padx=5, ipadx=5, ipady=5)
            
            drag_handle = ctk.CTkLabel(tile, text=" ☰ ", font=("Segoe UI", 20, "bold"), text_color="#7f8c8d", cursor="hand2")
            drag_handle.pack(side=tk.LEFT, fill=tk.Y, padx=(10, 15))
            drag_handle.bind("<Button-1>", lambda e, i=idx: self.obsluga_klikniecia_chwytaka(i))
            
            cf = ctk.CTkFrame(tile, fg_color="transparent")
            cf.pack(fill=tk.BOTH, expand=True)
            
            head = ctk.CTkFrame(cf, fg_color="transparent")
            head.pack(fill=tk.X, pady=(5, 5))
            
            if kategoria == "NAGŁÓWEK DNIA":
                ctk.CTkLabel(head, text=f"== {cw['nazwa']} ==", font=("Segoe UI", 14, "bold"), text_color="#e74c3c").pack(side=tk.LEFT)
                ctk.CTkButton(head, text="Usuń dzień X", font=("Segoe UI", 11, "bold"), fg_color="#c0392b", hover_color="#e74c3c", width=90, height=24, command=lambda i=idx: self.usun_konkretne_cwiczenie(i)).pack(side=tk.RIGHT, padx=5)
            else:
                ctk.CTkLabel(head, text=f"{idx + 1}. {kategoria.upper()}", font=("Segoe UI", 12, "bold"), text_color="#7f8c8d").pack(side=tk.LEFT)
                
                if not is_gym:
                    ctk.CTkLabel(head, text=f" {cw['czas_min']} MIN ", font=("Segoe UI", 12, "bold"), fg_color="#7f8c8d", text_color="#0f0f0f", corner_radius=4).pack(side=tk.RIGHT, padx=10)
                
                ctk.CTkButton(head, text="Usuń X", font=("Segoe UI", 11, "bold"), fg_color="#c0392b", hover_color="#e74c3c", width=60, height=24, command=lambda i=idx: self.usun_konkretne_cwiczenie(i)).pack(side=tk.RIGHT, padx=5)
                ctk.CTkButton(head, text="Edytuj ✎", font=("Segoe UI", 11, "bold"), fg_color="#d35400", hover_color="#e67e22", width=60, height=24, command=lambda i=idx: self.edytuj_parametry_w_planie(i)).pack(side=tk.RIGHT, padx=5)
            
            ctk.CTkLabel(cf, text=cw['nazwa'], font=("Segoe UI", 16, "bold"), text_color="#ffffff", anchor="w").pack(fill=tk.X, pady=(0, 2))
            ctk.CTkLabel(cf, text=f"➔ ZALECENIE: {cw['parametry']}", font=("Segoe UI", 14, "bold"), text_color="#2ecc71", anchor="w").pack(fill=tk.X, pady=(0, 2))
            
            if kategoria != "NAGŁÓWEK DNIA":
                ctk.CTkLabel(cf, text=f"[ANATOMIA: {cw['miesnie']}]", font=("Segoe UI", 12, "italic"), text_color="#94a3b8", anchor="w").pack(fill=tk.X, pady=(0, 4))
                
                uwagi_tekst = cw.get('uwagi', '')
                if uwagi_tekst:
                    ctk.CTkLabel(cf, text=f"📝 UWAGI: {uwagi_tekst}", font=("Segoe UI", 12, "bold"), text_color="#f39c12", anchor="w").pack(fill=tk.X, pady=(0, 4))
            
            # Bezpieczne zawijanie tekstu za pomocą sztywnego ograniczenia
            lbl_desc = ctk.CTkLabel(cf, text=cw['opis'], font=("Segoe UI", 13), text_color="#bdc3c7", justify="left", anchor="w", wraplength=700)
            lbl_desc.pack(fill=tk.X, pady=(0, 5))

    # ==============================================================================
    # 5. GENEROWANIE PLIKÓW (DOCX ORAZ EXCEL Z UWAGAMI)
    # ==============================================================================
    def zapisz_do_docx(self):
        if not self.wylosowany_plan_cache: return
        
        is_gym = self.ostatni_wybrany_profil.startswith("GYM:")
        postfix = "ilosciowy" if is_gym else f"{self.realny_czas_cache}min"
        
        sciezka = filedialog.asksaveasfilename(
            initialfile=f"Plan_treningowy_{postfix}.docx", 
            defaultextension=".docx", filetypes=[("Dokument Word", "*.docx")]
        )
        if not sciezka: return

        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(10)
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after = Pt(6)  
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        for s in doc.sections: 
            s.top_margin = Inches(1.0)
            s.bottom_margin = Inches(1.0)
            s.left_margin = Inches(1.0)
            s.right_margin = Inches(1.0)
            
        h = doc.add_heading(level=1)
        hr = h.add_run("ZINTEGROWANA KARTA REHABILITACYJNO-TRENINGOWA")
        hr.font.name = 'Arial'
        hr.font.size = Pt(14)
        hr.font.bold = True
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_metoda_label = p.add_run("Metodyka programowania: ")
        run_metoda_label.bold = True
        
        if is_gym:
            p.add_run("Dwa Niezależne Silniki Fizjo/Gym  |  Tryb: Trening siłowy (liczba ćwiczeń)\n")
        else:
            p.add_run(f"Dwa Niezależne Silniki Fizjo/Gym  |  Czas sesji (łączny): {self.realny_czas_cache} min\n")
        
        p_line = doc.add_paragraph()
        p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_line.add_run("__________________________________________________________________________________________________")
        
        for idx, (kat, cw) in enumerate(self.wylosowany_plan_cache, 1):
            if kat == "NAGŁÓWEK DNIA":
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_h = p_head.add_run(f"\n{cw['nazwa']}\n")
                run_h.bold = True
                run_h.font.size = Pt(12)
                continue

            p_cw = doc.add_paragraph()
            p_cw.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            if is_gym:
                run_title = p_cw.add_run(f"{idx}. {kat.upper()}: {cw['nazwa']}\n")
            else:
                run_title = p_cw.add_run(f"{idx}. {kat.upper()}: {cw['nazwa']} ({cw['czas_min']} min)\n")
            
            run_title.bold = True
            run_title.font.size = Pt(11)
            
            p_cw.add_run("    ● DAWKOWANIE: ").bold = True
            p_cw.add_run(f"{cw['parametry']}\n")
            
            p_cw.add_run("    ● ANATOMIA (MIĘŚNIE): ").bold = True
            p_cw.add_run(f"{cw['miesnie']}\n")
            
            p_cw.add_run("    ● OPIS I INSTRUKCJA: ").bold = True
            p_cw.add_run(f"{cw['opis']}\n")
            
            p_cw.add_run("    ● UWAGI: ").bold = True
            uwagi_text = cw.get('uwagi', '').strip()
            if uwagi_text:
                p_cw.add_run(f"{uwagi_text}\n")
            else:
                p_cw.add_run("............................................................................................................\n")

        try: 
            doc.save(sciezka)
            messagebox.showinfo("Sukces", "Karta treningowa (DOCX) została pomyślnie wyeksportowana.")
        except Exception as e: 
            messagebox.showerror("Błąd", str(e))

    def zapisz_do_excel(self):
        if not self.wylosowany_plan_cache: return
        
        try:
            liczba_dni = int(self.entry_dni.get())
            if liczba_dni < 1: raise ValueError
        except ValueError:
            messagebox.showerror("Błąd", "Wprowadź prawidłową (dodatnią) liczbę dni w oknie konfiguracji.")
            return

        sciezka = filedialog.asksaveasfilename(
            initialfile=f"Harmonogram_tygodniowy.xlsx", 
            defaultextension=".xlsx", 
            filetypes=[("Arkusz Excel", "*.xlsx")]
        )
        if not sciezka: return

        try:
            dane_excel = []
            is_gym = self.ostatni_wybrany_profil.startswith("GYM:")
            
            dane_excel.append(["Imię i Nazwisko:", "", "", "", "", "", ""])
            dane_excel.append(["Płeć:", "", "", "", "", "", ""])
            
            if is_gym:
                dane_excel.append(["Podsumowanie planu:", "Trening Siłowy (Ilościowy)", "", "", "", "", ""])
            else:
                dane_excel.append(["Całkowity czas planu:", f"{self.realny_czas_cache} min", "", "", "", "", ""])
                
            dane_excel.append(["", "", "", "", "", "", ""])
            
            czy_split = any(kat == "NAGŁÓWEK DNIA" for kat, cw in self.wylosowany_plan_cache)
            
            if czy_split:
                czasy_dni = []
                for k, c in self.wylosowany_plan_cache:
                    if k == "NAGŁÓWEK DNIA":
                        czasy_dni.append(0)
                    else:
                        if czasy_dni:
                            czasy_dni[-1] += c.get('czas_min', 0)

                lp = 1
                dzien_aktualny = 0
                for idx, (kat, cw) in enumerate(self.wylosowany_plan_cache):
                    if kat == "NAGŁÓWEK DNIA":
                        if dzien_aktualny > 0:
                            dane_excel.append(["", "", "", "", "", "", ""])
                        dzien_aktualny += 1
                        nazwa_dnia_i_partii = cw['nazwa']
                        
                        if is_gym:
                            dane_excel.append([f"{nazwa_dnia_i_partii}", "", "", "", "", "", ""])
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
                        uwagi_str = cw.get('uwagi', '')
                        dane_excel.append([lp, cw['nazwa'], czas_str, serie, powt, cw['opis'], uwagi_str])
                        lp += 1
                dane_excel.append(["", "", "", "", "", "", ""])
                
            else:
                for dzien in range(1, liczba_dni + 1):
                    if is_gym:
                        dane_excel.append([f"DZIEŃ {dzien}", "", "", "", "", "", ""])
                    else:
                        dane_excel.append([f"DZIEŃ {dzien} - Czas trwania: {self.realny_czas_cache} min", "", "", "", "", "", ""])
                        
                    dane_excel.append(["L.p.", "Nazwa ćwiczenia", "Czas", "Ilość serii", "Ilość powtórzeń", "Sposób wykonania", "Uwagi"])
                    
                    for idx, (kat, cw) in enumerate(self.wylosowany_plan_cache, 1):
                        p_lower = str(cw['parametry']).lower()
                        if 'x' in p_lower:
                            parts = p_lower.split('x')
                            serie = parts[0].replace('serie','').replace('seria','').strip()
                            powt = parts[1].replace('powtórzeń','').replace('powtórzenia','').strip()
                        else:
                            serie = "1"
                            powt = cw['parametry'].strip()

                        czas_str = "-" if is_gym else f"{cw.get('czas_min', 0)} min"
                        uwagi_str = cw.get('uwagi', '')
                        dane_excel.append([idx, cw['nazwa'], czas_str, serie, powt, cw['opis'], uwagi_str])
                    
                    dane_excel.append(["", "", "", "", "", "", ""])

            df = pd.DataFrame(dane_excel)
            writer = pd.ExcelWriter(sciezka, engine='openpyxl')
            df.to_excel(writer, index=False, header=False, sheet_name='Harmonogram')
            
            worksheet = writer.sheets['Harmonogram']
            worksheet.column_dimensions['A'].width = 8   
            worksheet.column_dimensions['B'].width = 30  
            worksheet.column_dimensions['C'].width = 10  
            worksheet.column_dimensions['D'].width = 12  
            worksheet.column_dimensions['E'].width = 20  
            worksheet.column_dimensions['F'].width = 50  
            worksheet.column_dimensions['G'].width = 25  
            
            bold_font = Font(bold=True)
            white_bold_font = Font(bold=True, color="FFFFFF")
            wrap_text = Alignment(wrap_text=True, vertical='top')
            center_align = Alignment(horizontal='center', vertical='center')
            
            thin_border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            
            fill_day = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")       
            fill_header = PatternFill(start_color="8DB4E2", end_color="8DB4E2", fill_type="solid")    
            fill_data_even = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid") 

            data_row_counter = 0

            for row_idx, row_data in enumerate(dane_excel, 1):
                val_a = str(row_data[0])
                
                if row_idx in [1, 2, 3]:
                    worksheet.cell(row=row_idx, column=1).font = bold_font
                    if row_idx in [1, 2]:
                        worksheet.cell(row=row_idx, column=2).border = Border(bottom=Side(style='thin'))
                        worksheet.cell(row=row_idx, column=3).border = Border(bottom=Side(style='thin'))
                        worksheet.cell(row=row_idx, column=4).border = Border(bottom=Side(style='thin'))
                    
                elif val_a.startswith("DZIEŃ") or val_a.startswith("Poniedziałek") or \
                     val_a.startswith("Wtorek") or val_a.startswith("Środa") or \
                     val_a.startswith("Czwartek") or val_a.startswith("Piątek") or \
                     val_a.startswith("Sobota") or val_a.startswith("Niedziela"):
                    worksheet.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=7)
                    cell = worksheet.cell(row=row_idx, column=1)
                    cell.font = white_bold_font
                    cell.fill = fill_day
                    cell.alignment = center_align
                    
                    for col_idx in range(1, 8):
                        worksheet.cell(row=row_idx, column=col_idx).border = thin_border
                    data_row_counter = 0
                    
                elif val_a == "L.p.":
                    for col_idx in range(1, 8):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        cell.font = bold_font
                        cell.fill = fill_header
                        cell.border = thin_border
                        cell.alignment = center_align
                        
                elif val_a.isdigit():
                    data_row_counter += 1
                    current_fill = fill_data_even if data_row_counter % 2 == 0 else None
                    
                    for col_idx in range(1, 8):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        cell.border = thin_border
                        if current_fill:
                            cell.fill = current_fill
                            
                        if col_idx in [6, 7]:
                            cell.alignment = wrap_text
                        elif col_idx in [1, 3, 4, 5]:
                            cell.alignment = center_align
            
            writer.close()
            messagebox.showinfo("Sukces", "Harmonogram treningowy został poprawnie zapisany i estetycznie ostylowany w formacie Excel.")
            
        except ImportError:
            messagebox.showerror("Brak biblioteki", "Brakuje silnika 'openpyxl'. Wpisz w terminalu: pip install openpyxl")
        except Exception as e:
            messagebox.showerror("Błąd", str(e))

    # ==============================================================================
    # ASYSTENT AI - GROQ (Szybki silnik Llama 3)
    # ==============================================================================
    def otworz_czat_ai(self):
        if hasattr(self, 'okno_czatu') and self.okno_czatu.winfo_exists():
            self.okno_czatu.focus()
            return

        self.okno_czatu = ctk.CTkToplevel(self.root)
        self.okno_czatu.title("Asystent AI - Groq (Llama 3)")
        self.okno_czatu.geometry("600x700")
        self.okno_czatu.configure(fg_color="#141414")
        self.okno_czatu.attributes('-topmost', True)

        # Inicjalizacja historii wiadomości w formacie wymaganym przez Groq
        self.historia_wiadomosci = [
            {"role": "system", "content": "Jesteś wirtualnym asystentem w aplikacji dla fizjoterapeutów i trenerów 'Fizjo Workout Ultimate'. Pomagasz profesjonalnie i zwięźle. Język: polski."}
        ]

        ctk.CTkLabel(self.okno_czatu, text="✨ Wirtualny Konsultant Treningowy (Groq)", font=("Segoe UI", 18, "bold"), text_color="#8e44ad").pack(pady=15)

        self.pole_historii = ctk.CTkTextbox(self.okno_czatu, font=("Segoe UI", 13), corner_radius=10, fg_color="#2b2b2b", text_color="#ffffff", wrap="word")
        self.pole_historii.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
        
        self.dopisz_do_czatu("AI: Cześć! Jestem Twoim wirtualnym asystentem napędzanym ultraszybkim silnikiem Groq. W czym mogę pomóc?\n")

        input_frame = ctk.CTkFrame(self.okno_czatu, fg_color="transparent")
        input_frame.pack(fill=tk.X, padx=20, pady=(0, 20))

        self.entry_czat = ctk.CTkEntry(input_frame, font=("Segoe UI", 13), placeholder_text="Wpisz swoje pytanie...", height=40)
        self.entry_czat.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.entry_czat.bind("<Return>", lambda event: self.wyslij_wiadomosc_ai())

        self.btn_wyslij_ai = ctk.CTkButton(input_frame, text="Wyślij", font=("Segoe UI", 12, "bold"), width=80, height=40, fg_color="#2ecc71", hover_color="#27ae60", text_color="#0f0f0f", command=self.wyslij_wiadomosc_ai)
        self.btn_wyslij_ai.pack(side=tk.RIGHT)

    def dopisz_do_czatu(self, tekst):
        self.pole_historii.configure(state="normal")
        self.pole_historii.insert(tk.END, tekst + "\n\n")
        self.pole_historii.configure(state="disabled")
        self.pole_historii.yview(tk.END) 

    def wyslij_wiadomosc_ai(self):
        pytanie = self.entry_czat.get().strip()
        if not pytanie:
            return

        self.dopisz_do_czatu(f"TY: {pytanie}")
        self.entry_czat.delete(0, tk.END)

        # Dodanie pytania do globalnej historii konwersacji (pamięć)
        self.historia_wiadomosci.append({"role": "user", "content": pytanie})

        self.btn_wyslij_ai.configure(state="disabled", text="Myśli...")
        self.entry_czat.configure(state="disabled")

        import threading
        threading.Thread(target=self.zapytanie_w_tle, daemon=True).start()

    def zapytanie_w_tle(self):
        try:
            # Wysłanie zapytania do API Groq z całą historią, aby model pamiętał kontekst
            completion = groq_client.chat.completions.create(
                model="llama-3.3-70b-versatile", # <--- ZAKTUALIZOWANA NAZWA MODELU
                messages=self.historia_wiadomosci,
                temperature=0.7,
                max_tokens=1024,
                top_p=1,
                stream=False
            )
            
            odpowiedz_tekst = completion.choices[0].message.content
            
            # Dodanie odpowiedzi AI do historii, by pamiętało ją w kolejnych pytaniach
            self.historia_wiadomosci.append({"role": "assistant", "content": odpowiedz_tekst})
            
            self.root.after(0, self.obsluga_odpowiedzi_ai, f"AI: {odpowiedz_tekst}")
        except Exception as e:
            # Nawet jeśli wystąpi błąd, usuwamy ostatnie pytanie z historii, aby nie zablokować czatu
            if self.historia_wiadomosci and self.historia_wiadomosci[-1]["role"] == "user":
                self.historia_wiadomosci.pop()
            self.root.after(0, self.obsluga_odpowiedzi_ai, f"BŁĄD API GROQ: {str(e)}")

    def obsluga_odpowiedzi_ai(self, tekst_odpowiedzi):
        self.dopisz_do_czatu(tekst_odpowiedzi)
        self.btn_wyslij_ai.configure(state="normal", text="Wyślij")
        self.entry_czat.configure(state="normal")
        self.entry_czat.focus()

if __name__ == '__main__':
    root = ctk.CTk()
    app = FizjoWorkoutUltimateApp(root)
    root.mainloop()