# ☁ Synchronizacja na żywo (Supabase) — instrukcja

Dzięki temu Twoje treningi synchronizują się **na żywo między telefonem a komputerem**,
a każdy kumpel ma **własne konto** na tym samym linku. Konfiguracja zajmuje ~10 minut
i jest **darmowa** na start.

> Dopóki tego nie ustawisz, aplikacja działa w 100% lokalnie (bez logowania).

---

## Krok 1 — Załóż projekt Supabase (darmowy)
1. Wejdź na **https://supabase.com** → **Start your project** → zaloguj się (GitHub/Google).
2. **New project**:
   - **Name:** `trening-analytics`
   - **Database Password:** wymyśl i zapisz (nie będzie potrzebne w apce)
   - **Region:** `Central EU (Frankfurt)` (najbliżej Polski)
3. Poczekaj ~2 min, aż projekt się postawi.

## Krok 2 — Utwórz bazę i magazyn (1 wklejka)
1. W panelu projektu: menu po lewej → **SQL Editor** → **New query**.
2. Otwórz plik **`supabase-schema.sql`** (z tego repo), skopiuj całość, wklej i kliknij **Run**.
3. Powinno wyświetlić „Success”. To tworzy tabelę, prywatny magazyn plików, RLS i realtime.

## Krok 3 — Ustawienia logowania
1. Menu → **Authentication** → **Providers** → upewnij się, że **Email** jest włączony.
2. (Opcjonalnie, dla wygody kumpli) **Authentication → Sign In / Providers → Email**:
   wyłącz **„Confirm email”**, jeśli nie chcesz potwierdzania mailem przy rejestracji.
   (Z włączonym potwierdzaniem każdy musi kliknąć link z maila przed pierwszym logowaniem.)

## Krok 4 — Wklej klucze do aplikacji
1. Menu → **Project Settings** (⚙) → **API**.
2. Skopiuj:
   - **Project URL** (np. `https://abcdefgh.supabase.co`)
   - **anon public** (długi klucz `eyJ...`) — jest bezpieczny do publikacji, chroni go RLS.
3. Otwórz **`app/config.js`** i wpisz te dwie wartości:
   ```js
   window.TA_CONFIG = {
     SUPABASE_URL: 'https://abcdefgh.supabase.co',
     SUPABASE_ANON_KEY: 'eyJhbGciOi...'
   };
   ```
4. Zapisz, wtypchnij na GitHub (`git add app/config.js && git commit -m "config chmury" && git push`)
   — lub po prostu prześlij mi te 2 wartości, wkleję i wypchnę.

## Krok 5 — Gotowe ✅
- W aplikacji pojawi się przycisk **☁ Zaloguj** (prawy górny róg).
- Załóż konto e-mailem → od teraz każdy zapis/eksport ląduje w chmurze i **synchronizuje się
  na żywo** na każdym zalogowanym urządzeniu (kropka przy ☁: niebieska = synchronizuję,
  zielona = zapisane).

---

## 📱 Instalacja jako apka na Androidzie (Xiaomi 14T Pro)
1. Otwórz link do aplikacji w **Chrome**.
2. Menu (⋮) → **Dodaj do ekranu głównego** / **Zainstaluj aplikację**.
3. Ikona pojawi się jak normalna apka (pełny ekran, działa też offline do przeglądania).

## 👥 Dla kumpli
- Wystarczy, że wejdą na **ten sam link** i klikną **☁ Zaloguj → Załóż konto**.
- Każdy ma swoje, oddzielone dane (RLS). Nikt nie widzi cudzych treningów.

## Koszty
- Supabase **Free**: w zupełności wystarczy dla Ciebie i paczki znajomych.
- Jeśli kiedyś urośnie — **Pro ok. $25/mc**. Hosting strony (GitHub Pages) jest darmowy.
