/* =====================================================================
   cloud.js — synchronizacja na żywo + konta (Supabase)
   ---------------------------------------------------------------------
   Model: w chmurze trzymamy CAŁY plik .xlsx użytkownika (zachowuje
   formatowanie i obieg z trenerem) + wiersz stanu z czasem aktualizacji.
   Zapis na jednym urządzeniu → realtime powiadamia drugie → ono pobiera
   nowy plik i odświeża wykresy. Dane chroni RLS (każdy widzi tylko swoje).

   Gdy brak konfiguracji (config.js puste) — moduł jest wyłączony, a apka
   działa lokalnie jak dotychczas.
   ===================================================================== */
window.Cloud = (function () {
  'use strict';
  const BUCKET = 'workbooks';
  let client = null;
  let enabled = false;
  let user = null;
  let channel = null;
  const authCbs = [];
  let remoteCb = null;

  // identyfikator urządzenia (żeby ignorować własne echa realtime)
  function deviceId() {
    let d = localStorage.getItem('ta_device');
    if (!d) { d = 'dev_' + Math.random().toString(36).slice(2, 10); localStorage.setItem('ta_device', d); }
    return d;
  }

  function init() {
    const cfg = window.TA_CONFIG || {};
    if (!cfg.SUPABASE_URL || !cfg.SUPABASE_ANON_KEY || !window.supabase) {
      enabled = false; return false;
    }
    client = window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY, {
      auth: { persistSession: true, autoRefreshToken: true }
    });
    enabled = true;
    client.auth.onAuthStateChange((_event, session) => {
      user = session ? session.user : null;
      if (user) startRealtime(); else stopRealtime();
      authCbs.forEach((cb) => { try { cb(user); } catch (e) {} });
    });
    // odczytaj istniejącą sesję
    client.auth.getSession().then(({ data }) => {
      user = (data && data.session) ? data.session.user : null;
      if (user) startRealtime();
      authCbs.forEach((cb) => { try { cb(user); } catch (e) {} });
    }).catch((e) => { console.warn('Cloud: brak sesji', e); });
    return true;
  }

  const isEnabled = () => enabled;
  const getUser = () => user;
  const onAuth = (cb) => { authCbs.push(cb); };

  async function signUp(email, password) {
    const { data, error } = await client.auth.signUp({ email, password });
    if (error) throw error;
    return data;
  }
  async function signIn(email, password) {
    const { data, error } = await client.auth.signInWithPassword({ email, password });
    if (error) throw error;
    return data;
  }
  async function signOut() { await client.auth.signOut(); user = null; }

  // wyślij plik do chmury + zaktualizuj stan
  async function pushWorkbook(bytes, meta) {
    if (!enabled || !user) return;
    const path = user.id + '/current.xlsx';
    const blob = bytes instanceof Blob ? bytes
      : new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const up = await client.storage.from(BUCKET).upload(path, blob, {
      upsert: true, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    if (up.error) throw up.error;
    const row = {
      user_id: user.id,
      updated_at: new Date().toISOString(),
      file_name: (meta && meta.fileName) || 'trening.xlsx',
      week_count: (meta && meta.weekCount) || null,
      device: deviceId()
    };
    const { error } = await client.from('workbook_state').upsert(row, { onConflict: 'user_id' });
    if (error) throw error;
  }

  // pobierz najnowszy plik z chmury (lub null)
  async function pullWorkbook() {
    if (!enabled || !user) return null;
    const { data: st } = await client.from('workbook_state').select('*').eq('user_id', user.id).maybeSingle();
    if (!st) return null;
    const dl = await client.storage.from(BUCKET).download(user.id + '/current.xlsx');
    if (dl.error || !dl.data) return null;
    const buf = await dl.data.arrayBuffer();
    return { bytes: new Uint8Array(buf), meta: st };
  }

  function startRealtime() {
    stopRealtime();
    if (!user) return;
    channel = client.channel('wb-' + user.id)
      .on('postgres_changes',
        { event: '*', schema: 'public', table: 'workbook_state', filter: 'user_id=eq.' + user.id },
        (payload) => {
          const row = payload.new || {};
          if (row.device && row.device === deviceId()) return; // własna zmiana
          if (remoteCb) remoteCb(row);
        })
      .subscribe();
  }
  function stopRealtime() { if (channel) { try { client.removeChannel(channel); } catch (e) {} channel = null; } }
  function onRemoteChange(cb) { remoteCb = cb; }

  return { init, isEnabled, getUser, onAuth, signUp, signIn, signOut, pushWorkbook, pullWorkbook, onRemoteChange };
})();
