/* =====================================================================
   Konfiguracja chmury (Supabase) — synchronizacja na żywo + konta.
   ---------------------------------------------------------------------
   Wpisz tu 2 wartości ze swojego projektu Supabase:
     Project URL   (Settings → API → Project URL)
     anon public   (Settings → API → Project API keys → anon public)
   Klucz „anon public” jest BEZPIECZNY do publikacji — dostęp do danych
   chroni RLS (Row Level Security), które ustawiamy w bazie.

   Dopóki te pola są puste, aplikacja działa w 100% lokalnie (bez chmury).
   Pełna instrukcja: SETUP-CHMURA.md
   ===================================================================== */
window.TA_CONFIG = {
  SUPABASE_URL: '',      // np. 'https://abcdefgh.supabase.co'
  SUPABASE_ANON_KEY: ''  // np. 'eyJhbGciOi...'
};
