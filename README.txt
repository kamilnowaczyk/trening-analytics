════════════════════════════════════════════════════════════════════
  TRENING ANALYTICS — Kamil
  Lokalna aplikacja do analizy raportów treningowych (siła + sylwetka)
════════════════════════════════════════════════════════════════════

JAK URUCHOMIĆ
─────────────
Kliknij dwukrotnie plik:  „Trening Analytics.html”
Otworzy się w przeglądarce (najlepiej Chrome lub Edge). Internet NIE jest
potrzebny — wszystko działa lokalnie.


CO TYDZIEŃ (Twój obieg pracy)
─────────────────────────────
1. Trener przysyła nowy plik .xlsx (o tydzień dłuższy, z progresją).
2. Otwierasz aplikację → „Otwórz raport” → wskazujesz najnowszy plik.
3. Aplikacja od razu pokazuje wszystkie wykresy i statystyki.
   (Czyta WSZYSTKIE tygodnie — także te w kolumnach, które trener ukrył.)
4. Zakładka „Wpis tygodnia” → wpisujesz wykonane ciężary i serie z treningu,
   a pod każdą serią wybierasz KOLOR wysiłku (zielony/pomarańczowy/czerwony)
   — dokładnie jak w pliku.
5. „Zapisz do pliku” → dane trafiają z powrotem do tego samego .xlsx,
   z zachowaniem formatowania i KOLORÓW (użyte są istniejące style z Twojego
   pliku, więc wygląda 1:1). Albo „Eksportuj dla trenera” → pobiera kopię
   w tej samej formie, z nazwą zawierającą numer tygodnia. Ten plik wysyłasz.


ZAKŁADKI
────────
• Pulpit            — przegląd: waga, siła, tonaż, kroki, sen.
• Siła & progresja  — dla KAŻDEGO ćwiczenia: ciężar, szac. 1RM, tonaż,
                      tabela serii, rekordy (★ = nowy rekord 1RM).
• Sylwetka & pomiary— waga, obwody (udo/pas/biodra/klatka/biceps),
                      kalorie, białko, sen, kroki.
• Raporty tygodniowe— Twoje cotygodniowe odpowiedzi + ocena planu i głód.
• Wpis tygodnia     — uzupełnianie i zapis danych do pliku.
• Skala RPE         — ściąga (zapas powtórzeń) + legenda kolorów.


ZAPIS DO PLIKU — dwa tryby
──────────────────────────
• Chrome / Edge, gdy otworzysz plik przyciskiem „Otwórz raport”:
  aplikacja zapisuje wpisy BEZPOŚREDNIO w tym samym pliku (w miejscu).
  Aplikacja zapamiętuje plik — przy następnym otwarciu sama go wczyta
  (wystarczy potwierdzić dostęp).

• Inne przeglądarki / wczytanie przez przeciągnięcie:
  „Zapisz” pobiera nową kopię pliku „...(edytowany).xlsx” do folderu
  Pobrane — i ten plik wysyłasz trenerowi.

Wskazówka (pewny zapis w miejscu): zamiast dwukliku możesz uruchomić
folder przez lokalny serwer, np. w tym katalogu:
    python -m http.server 8765
a potem wejść na  http://localhost:8765/Trening%20Analytics.html


STRUKTURA PLIKÓW (nie kasuj podfolderów)
────────────────────────────────────────
  Trening Analytics.html   ← uruchamiasz to
  app/   — wygląd i logika aplikacji
  lib/   — biblioteki (działają offline) + silnik czytania/zapisu .xlsx
  dane/  — (opcjonalnie) miejsce na pliki .xlsx


KOLORY SERII, TOP-SETY I ZMIANY ĆWICZEŃ
───────────────────────────────────────
• Kolory (RIR): zielony = zapas 2–4 powt., pomarańczowy = zapas 1–2,
  czerwony = seria nieudana. Aplikacja CZYTA te kolory z pliku (pokazuje
  je w tabeli „Siła”) i pozwala je USTAWIĆ przy wpisie — zapis używa
  Twoich istniejących stylów, więc Excel renderuje identycznie.
• Top-sety (odchył od reguły): wpisy typu „110(6)”, „130kg/6”, „115(7)”
  są rozpoznawane jako jedna cięższa seria (np. 110 kg na 6), a pozostałe
  serie liczone na bazowym obciążeniu. W tabeli oznaczone ▲. Tonaż i
  szac. 1RM uwzględniają realny ciężar każdej serii.
• Zmiana ćwiczenia: plik nie zapisuje historycznych nazw ćwiczeń —
  każdy wiersz to „slot” w planie, a nazwa jest aktualna. Gdy zmienia się
  typ obciążenia w danym wierszu (sztanga/hantle/maszyna), aplikacja
  pokazuje ostrzeżenie „możliwa zmiana ćwiczenia”, żeby nie mylić
  starszych tygodni z poprzednim ćwiczeniem.

JAK TO DZIAŁA Z UKRYTYMI KOLUMNAMI
──────────────────────────────────
Plik .xlsx to spakowane dane XML. Aplikacja czyta wartości komórek po
ich adresie (np. DI5), a tygodnie rozpoznaje po POZYCJI kolumn
(tydzień N = kolumna 17 + (N-1)×5), nie po nagłówkach. Ukrycie kolumny
w Excelu to tylko wygląd — dane są zawsze obecne, więc nic nie ginie,
nawet gdy trener raz po raz chowa inne kolumny.
