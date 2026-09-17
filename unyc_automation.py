"""
UNYC Atlas Automation Script (robuste)
- Login (sans secrets en dur)
- Navigation client
- Ouverture "Lignes mobiles" -> ligne -> onglet Contrat
- Lecture RIO par input.value
- Écriture RIO sur la ligne Excel (Client + Numéro)

Dépendances:
  pip install playwright pandas
  playwright install
"""

import asyncio
import os
import re
from dataclasses import dataclass,field
from typing import Optional, List, Dict
from pandas import DataFrame
from getpass import getpass
from contextlib import suppress

import pandas as pd
from playwright.async_api import (
    async_playwright,
    Page,
    Browser,
    BrowserContext,
    TimeoutError as PlaywrightTimeoutError,
)
import os, pathlib
print("CWD:", pathlib.Path().resolve())
print(".env exists:", pathlib.Path(".env").exists())
print("ENV USER set?", bool(os.getenv("UNYC_USERNAME")))
print("ENV PASS set?", bool(os.getenv("UNYC_PASSWORD")))

# =========================
# Configuration & Helpers
# =========================

@dataclass
class Settings:
    username: str = field(default_factory=lambda: os.getenv("UNYC_USERNAME", ""))
    password: str = field(default_factory=lambda: os.getenv("UNYC_PASSWORD", ""))
    base_url: str = field(default_factory=lambda: os.getenv("UNYC_BASE_URL", "https://atlas.unyc.io/"))
    client_url: str = field(default_factory=lambda: os.getenv("UNYC_CLIENT_URL", "https://atlas.unyc.io/client"))
    excel_file: str = field(default_factory=lambda: os.getenv("UNYC_EXCEL", "lmunyc.xlsx"))
    headless: bool = field(default_factory=lambda: os.getenv("UNYC_HEADLESS", "false").lower() in ("1", "true", "yes"))
    slow_mo_ms: int = field(default_factory=lambda: int(os.getenv("UNYC_SLOW_MO_MS", "200")))
    storage_state_path: Optional[str] = field(default_factory=lambda: os.getenv("UNYC_STORAGE_STATE", "unyc_state.json"))



from getpass import getpass

class UnycAutomation:
    def __init__(self):
        self.settings = Settings()
        # Fallback interactif si non défini
        if not self.settings.username:
            self.settings.username = input("UNYC username: ").strip()
        if not self.settings.password:
            self.settings.password = getpass("UNYC password: ")
        self.processed_clients: List[Dict] = []
        self.excel_df: Optional[DataFrame] = None


    # --------- Normalisation numéros ---------
    def _normalize_phone(self, v: str) -> str:
        s = str(v or "").strip()
        if s.endswith(".0"):  # Excel qui convertit en float
            s = s[:-2]
        if s.startswith("+"):
            return "+" + "".join(ch for ch in s[1:] if ch.isdigit())
        return "".join(ch for ch in s if ch.isdigit())

    # =========================
    # Excel
    # =========================
    def read_clients_from_excel(self) -> List[str]:
        """Charge l'Excel, prépare colonnes et normalisation, retourne la liste unique de clients."""
        try:
            if not os.path.exists(self.settings.excel_file):
                print(f"❌ Fichier Excel '{self.settings.excel_file}' introuvable.")
                return []

            print(f"📊 Lecture: {self.settings.excel_file} (Feuil1)")
            df = pd.read_excel(self.settings.excel_file, sheet_name="Feuil1", dtype=str)

            if "Client" not in df.columns:
                print("❌ Colonne 'Client' introuvable.")
                print(f"Colonnes disponibles: {list(df.columns)}")
                return []

            if "RIO" not in df.columns:
                df["RIO"] = ""
                print("✅ Ajout de la colonne 'RIO'")

            if "Numéro" not in df.columns:
                df["Numéro"] = ""
                print("⚠️ Colonne 'Numéro' absente — ajoute des numéros pour traiter les RIO")

            df["__NormNumero"] = df["Numéro"].apply(self._normalize_phone)
            self.excel_df = df

            clients = (
                df["Client"].fillna("").map(str.strip)
            )
            clients = sorted(set([c for c in clients if c]))
            print(f"✅ {len(clients)} client(s) à traiter")
            return clients

        except Exception as e:
            print(f"❌ Erreur lecture Excel: {e}")
            return []
    def _df(self) -> DataFrame:
        """Retourne le DataFrame Excel chargé, sinon lève une erreur claire."""
        if self.excel_df is None:
            raise RuntimeError(
                "Excel non chargé. Appelle d'abord read_clients_from_excel() et vérifie qu'il a réussi."
            )
        return self.excel_df
    def get_phone_numbers_for_client(self, client_name: str) -> List[str]:
        """Retourne les numéros (tels qu'en Excel) pour un client donné."""
        try:
            df = self._df()  
            rows = df[df["Client"].str.strip() == client_name.strip()]
            if rows.empty:
                print(f"❌ Client {client_name} introuvable dans l'Excel")
                return []
            phones = [str(x).strip() for x in rows["Numéro"].tolist() if str(x or "").strip()]
            print(f"📞 {len(phones)} numéro(s) pour {client_name}: {phones}")
            return phones
        except Exception as e:
            print(f"❌ Erreur extraction numéros ({client_name}): {e}")
            return []


    def save_rio_to_excel(self, client_name: str, phone_number: str, rio_code: str) -> bool:
        """Écrit le RIO sur la ligne correspondant au couple (Client, Numéro)."""
        try:
            df = self._df()  
            norm = self._normalize_phone(phone_number)
            mask = (
                df["Client"].str.strip().eq(client_name.strip())
                & df["__NormNumero"].eq(norm)
            )
            if mask.any():
                df.loc[mask, "RIO"] = rio_code
                # On crée une COPIE pour la sauvegarde sans la colonne technique
                save_df = df.drop(columns=["__NormNumero"]) if "__NormNumero" in df.columns else df
                save_df.to_excel(self.settings.excel_file, sheet_name="Feuil1", index=False)
                print(f"✅ RIO {rio_code} enregistré pour {client_name} / {phone_number}")
                return True

            print(f"❌ Aucune ligne correspondante pour {client_name} / {phone_number}")
            return False

        except Exception as e:
            print(f"❌ Erreur écriture RIO: {e}")
            return False


    # =========================
    # Playwright
    # =========================
    async def setup_browser(self, playwright):
        """Crée browser + context, réutilise éventuellement un storage_state."""
        browser = await playwright.chromium.launch(
            headless=self.settings.headless,
            slow_mo=self.settings.slow_mo_ms,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
            ],
        )

        # Si on a un storage_state existant, l'utiliser
        context_kwargs = dict(
            viewport={"width": 1600, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            java_script_enabled=True,
            accept_downloads=True,
        )
        storage_path = self.settings.storage_state_path
        if storage_path and os.path.exists(storage_path):
            context_kwargs["storage_state"] = storage_path

        context = await browser.new_context(**context_kwargs)
        page = await context.new_page()
        return browser, context, page
    async def _post_login_visible(self, page: Page, timeout: int = 4000) -> bool:
        """
        Renvoie True si on voit des marqueurs d'une page "connectée"
        (DOM interne ou URL), sinon False.
        """
        marker = page.locator('#customer_list_search, #telephonie_client_panel, nav, header').first
        try:
            await marker.wait_for(state='visible', timeout=timeout)
            return True
        except PlaywrightTimeoutError:
            href = page.url or ""
            return any(path in href for path in ("/client", "/accueil", "/revendeur"))

    async def login(self, context: BrowserContext, page: Page):
        """Login robuste + vérification de la page juste après, réutilise la session si dispo."""
        print("🌐 Ouverture page de login…")
        await page.goto(self.settings.base_url)

        # (A) Si storage_state a gardé la session, on est peut-être déjà connecté
        if await self._post_login_visible(page, timeout=2000):
            print(f"🔁 Session déjà active (storage_state) → {page.url}")
            # Sécurise: navigue vers /client si on n'y est pas
            if "/client" not in (page.url or ""):
                await page.goto(self.settings.client_url)
                if await self._post_login_visible(page, timeout=4000):
                    print(f"✅ Page chargée après réutilisation: {page.url}")
            # Resauvegarde l'état pour prolonger la session
            if self.settings.storage_state_path:
                with suppress(Exception):
                    await context.storage_state(path=self.settings.storage_state_path)
            return

        # (B) Sinon, on fait un vrai login
        await page.wait_for_selector('input[name="username"]', timeout=20_000)
        await page.fill('input[name="username"]', self.settings.username)
        await page.fill('input[name="password"]', self.settings.password)
        await page.click('button[type="submit"]')

        # -- Sélecteurs d'état --
        twofa = page.locator(
            'input[autocomplete="one-time-code"], input#first, input#second, input#third, '
            'input#fourth, input#fifth, input#sixth, input[type="number"][placeholder="_"]'
        )
        post_login = page.locator('#customer_list_search, #telephonie_client_panel, nav, header').first

        # -- Course entre "2FA" et "post-login" --
        task_twofa = asyncio.create_task(twofa.first.wait_for(state="visible"))
        task_post = asyncio.create_task(post_login.wait_for(state="visible"))

        done, pending = await asyncio.wait(
            {task_twofa, task_post},
            timeout=60,
            return_when=asyncio.FIRST_COMPLETED,
        )

        for t in pending:
            t.cancel()
            with suppress(asyncio.CancelledError):
                await t

        if task_twofa in done:
            print("🔐 2FA détectée. Saisis le code dans le navigateur (attente max 3 min)…")
            with suppress(PlaywrightTimeoutError):
                await post_login.wait_for(state="visible", timeout=180_000)

        # (C) Vérification explicite: quelle page est chargée juste après le login ?
        if await self._post_login_visible(page, timeout=10_000):
            print(f"✅ Page chargée après login: {page.url}")
        else:
            print("⚠️ Post-login non détecté. Tentative de navigation vers la page client…")
            await page.goto(self.settings.client_url)
            if await self._post_login_visible(page, timeout=10_000):
                print(f"✅ Page client accessible: {page.url}")
            else:
                print("❌ Impossible de confirmer l'état connecté.")

        # (D) Sauvegarder/rafraîchir le storage_state si demandé
        if self.settings.storage_state_path:
            try:
                await context.storage_state(path=self.settings.storage_state_path)
                print(f"💾 Session sauvegardée dans {self.settings.storage_state_path}")
            except Exception as e:
                print(f"⚠️ Impossible de sauvegarder la session: {e}")



    async def navigate_to_client_page(self, page: Page) -> bool:
        print("➡️ Navigation vers la page client…")
        await page.goto(self.settings.client_url)
        try:
            await page.wait_for_selector("input#customer_list_search", timeout=20_000)
            print("✅ Page client prête.")
            return True
        except PlaywrightTimeoutError:
            print("❌ Champ de recherche client introuvable.")
            return False

    async def search_and_click_client(self, page: Page, client_name: str) -> bool:
        """Recherche un client et ouvre sa fiche."""
        try:
            print(f"🔍 Recherche client: {client_name}")
            await page.fill("input#customer_list_search", "")
            await page.type("input#customer_list_search", client_name, delay=80)
            await page.wait_for_timeout(1200)
            await page.wait_for_selector("table#liste tbody", timeout=10_000)

            client_links = await page.query_selector_all('table#liste tbody tr.row-client td a[href*="/client/"]')
            if not client_links:
                print(f"❌ Aucun résultat pour: {client_name}")
                return False

            for link in client_links:
                text = (await link.inner_text()).strip()
                if client_name.lower() in text.lower():
                    href = await link.get_attribute("href")
                    print(f"✅ Client trouvé: {text} -> {href}")
                    await link.click()
                    await page.wait_for_load_state("networkidle", timeout=20_000)

                    self.processed_clients.append({
                        "client_name": client_name,
                        "found_text": text,
                        "url": page.url,
                        "status": "success",
                    })

                    await self.extract_rio_from_client_page(page, client_name)

                    # retour à la liste
                    await page.goto(self.settings.client_url)
                    await page.wait_for_selector("input#customer_list_search", timeout=20_000)
                    return True

            print(f"❌ Client '{client_name}' non trouvé dans les résultats")
            self.processed_clients.append({
                "client_name": client_name,
                "found_text": "",
                "url": "",
                "status": "not_found",
            })
            return False

        except Exception as e:
            print(f"❌ Erreur ouverture fiche client '{client_name}': {e}")
            self.processed_clients.append({
                "client_name": client_name,
                "found_text": "",
                "url": "",
                "status": f"error: {e}",
            })
            return False

    # =========================
    # Extraction RIO
    # =========================
    async def _read_rio_from_scope(self, page: Page, scope_locator, user_id: str) -> Optional[str]:
        """
        Tente d'extraire un RIO (12 alphanum en majuscules) dans le scope fourni :
        1) value / attributs de l'input
        2) Ctrl+A / Ctrl+C (même si l'input est disabled)
        3) Scan du texte visible dans le panneau "Contrat"
        """
        import re
        from contextlib import suppress

        RIO_RE = re.compile(r"\b[A-Z0-9]{12}\b")

        def _clean(s: str) -> str:
            return (s or "").replace(" ", "").upper()

        # Installe un "sniffer" de copier pour récupérer le texte copié sans permissions clipboard
        with suppress(Exception):
            await page.evaluate("""
                () => {
                if (window.__copySnifferInstalled) return;
                window.__lastCopied = "";
                document.addEventListener('copy', function(e){
                    let txt = "";
                    const a = document.activeElement;
                    try {
                    if (a && 'value' in a) {
                        const el = a;
                        const s = el.selectionStart || 0;
                        const epos = el.selectionEnd || 0;
                        txt = (el.value || '').substring(s, epos) || (el.value || '');
                    }
                    } catch(_){}
                    if (!txt) {
                    try { txt = (window.getSelection && window.getSelection().toString()) || ""; } catch(_){}
                    }
                    try { e.clipboardData.setData('text/plain', txt); } catch(_){}
                    window.__lastCopied = txt || "";
                    e.preventDefault();
                }, {capture:true});
                window.__copySnifferInstalled = true;
                }
            """)

        # Candidats dans le scope "Contrat"
        candidates = [
            f'#contract_mobile_{user_id} input#licence_mobile_edit_rio_{user_id}',
            f'#licence_mobile_{user_id} input#licence_mobile_edit_rio_{user_id}',
            f'input#licence_mobile_edit_rio_{user_id}',
            f'input[name="licence_mobile_edit_rio_{user_id}"]',
            'input[id*="rio"]',
        ]

        # On réessaie quelques fois au cas où le champ se remplit après XHR
        for attempt in range(10):
            for css in candidates:
                loc = scope_locator.locator(css)
                cnt = await loc.count()
                if cnt == 0:
                    continue

                for i in range(min(cnt, 3)):
                    el = loc.nth(i)
                    with suppress(Exception):
                        await el.scroll_into_view_if_needed()

                    # 1) lecture directe (input_value + attributs)
                    with suppress(Exception):
                        v = (await el.input_value()).strip()
                        if v and RIO_RE.fullmatch(_clean(v)):
                            return _clean(v)

                    with suppress(Exception):
                        h = await el.element_handle()
                        vals = await page.evaluate("""
                            el => {
                            return {
                                value: (el && el.value) || "",
                                attr: (el && el.getAttribute && el.getAttribute('value')) || "",
                                defv: (el && el.defaultValue) || "",
                                placeholder: (el && el.getAttribute && el.getAttribute('placeholder')) || "",
                                disabled: !!(el && el.hasAttribute && el.hasAttribute('disabled'))
                            };
                            }
                        """, h)
                        for raw in (vals.get("value",""), vals.get("attr",""), vals.get("defv",""), vals.get("placeholder","")):
                            if raw and RIO_RE.fullmatch(_clean(raw)):
                                return _clean(raw)

                    # 2) Ctrl+A / Ctrl+C (même si disabled)
                    with suppress(Exception):
                        h = await el.element_handle()
                        await page.evaluate("""
                            el => {
                            const was = el.hasAttribute('disabled');
                            if (was) el.removeAttribute('disabled');
                            try { el.focus(); } catch(_){}
                            try { el.select && el.select(); } catch(_){}
                            el.__wasDisabled = was;
                            }
                        """, h)
                        await page.keyboard.press("Control+a")
                        await page.keyboard.press("Control+c")
                        await page.wait_for_timeout(120)
                        copied = await page.evaluate("window.__lastCopied || ''")
                        if copied and RIO_RE.fullmatch(_clean(copied)):
                            with suppress(Exception):
                                await page.evaluate("el => { if (el.__wasDisabled) el.setAttribute('disabled',''); }", h)
                            return _clean(copied)
                        # remettre disabled si besoin
                        with suppress(Exception):
                            await page.evaluate("el => { if (el.__wasDisabled) el.setAttribute('disabled',''); }", h)

            # pas trouvé → petite attente et on retente (le temps que l’input se remplisse)
            await page.wait_for_timeout(400)

        # 3) Dernier filet: scanner le texte du panneau "Contrat" dans le scope
        with suppress(Exception):
            panel = scope_locator.locator(f"#contract_mobile_{user_id}")
            if await panel.count():
                txt = await panel.first.inner_text()
                if txt:
                    m = RIO_RE.search(_clean(txt))
                    if m:
                        return m.group(0)

        return None

    async def extract_rio_from_client_page(self, page: Page, client_name: str):
        """Depuis la fiche client ouverte, trouve les lignes/IDs et extrait les RIO pour les numéros présents en Excel."""
        try:
            print(f"📱 Démarrage extraction RIO pour: {client_name}")
            target_numbers = self.get_phone_numbers_for_client(client_name)
            if not target_numbers:
                print(f"❌ Aucun numéro Excel pour {client_name}")
                return

            await page.wait_for_selector("#telephonie_client_panel", timeout=20_000)

            # afficher "Tous" (si présent)
            try:
                select_selector = 'select[name="table_licences_utilisateur_length"]'
                await page.wait_for_selector(select_selector, timeout=5_000)
                await page.select_option(select_selector, value="-1")
                await page.wait_for_timeout(600)
            except PlaywrightTimeoutError:
                pass

            await page.wait_for_selector("table#table_licences_utilisateur", timeout=10_000)
            phone_rows = await page.query_selector_all("table#table_licences_utilisateur tbody tr")

            # Pour chaque numéro Excel, trouver la ligne/ID correspondante
            for target in target_numbers:
                print(f"\n🔎 Recherche du numéro cible: {target}")
                norm_target = self._normalize_phone(target)
                found = False

                for row in phone_rows:
                    number_cell = await row.query_selector("td.sorting_3")
                    if not number_cell:
                        continue
                    dom_text = (await number_cell.inner_text()).strip()
                    dom_norm = self._normalize_phone(dom_text)

                    if dom_norm == norm_target:
                        row_id = await row.get_attribute("id")
                        if row_id and "licence_row_utilisateur-" in row_id:
                            user_id = row_id.replace("licence_row_utilisateur-", "")
                            print(f"✅ Numéro trouvé: {dom_text} (ID: {user_id})")
                            await self.extract_rio_for_phone_number(page, client_name, {
                                "number": target,
                                "found_number": dom_text,
                                "user_id": user_id,
                                "row": row,
                            })
                            found = True
                            break

                if not found:
                    print(f"❌ Numéro {target} introuvable sur la page client")
                    self.processed_clients.append({
                        "client_name": client_name,
                        "phone_number": target,
                        "rio_code": "Not found on page",
                        "status": "not_found_on_page",
                    })

        except Exception as e:
            print(f"❌ Erreur extraction RIO ({client_name}): {e}")

    def _spaced_pattern(self, digits: str) -> re.Pattern:
        # "0756005665" -> r"0\s*7\s*5\s*6\s*0\s*0\s*5\s*6\s*6\s*5"
        parts = [re.escape(d) for d in digits]
        return re.compile(r"\s*".join(parts))

    async def _wait_any_selector(self, page: Page, selectors: List[str], timeout: int = 15000):
        """Attend qu'au moins un des selecteurs existe (et ne soit pas display:none)."""
        js = """
        sels => {
        for (const sel of sels) {
            const el = document.querySelector(sel);
            if (!el) continue;
            const cs = getComputedStyle(el);
            if (cs && cs.display !== 'none') return true;
        }
        return false;
        }
        """
        return await page.wait_for_function(js, arg=selectors, timeout=timeout)
    async def extract_rio_for_phone_number(self, page: Page, client_name: str, phone_info: Dict):
        """Ouvre la sous-ligne, va sur l'onglet Contrat, lit le RIO et l'enregistre."""
        target_number = phone_info.get("number", "Unknown")
        found_number = phone_info.get("found_number", target_number)
        user_id = phone_info.get("user_id", "")
        row = phone_info.get("row")

        async def wait_any(selectors: List[str], timeout: int = 15000):
            js = """
            sels => {
            for (const sel of sels) {
                const el = document.querySelector(sel);
                if (!el) continue;
                const cs = getComputedStyle(el);
                if (!cs) continue;
                if (el.offsetParent !== null || cs.display !== 'none' || cs.visibility !== 'hidden') return true;
            }
            return false;
            }
            """
            return await page.wait_for_function(js, arg=selectors, timeout=timeout)

        try:
            print(f"📂 Traitement: {target_number} (affiché: {found_number})")

            if not row:
                print(f"❌ Pas d'élément de ligne pour {target_number}")
                return

            # Étendre la ligne (icône détails)
            details_control = await row.query_selector("td.details-control")
            if not details_control:
                print(f"❌ Icône détails introuvable pour {target_number}")
                return

            # clic robuste
            try:
                await details_control.click()
            except Exception:
                await page.evaluate("(el)=>el.click()", details_control)

            # Attente robuste d'au moins un marqueur de la fiche détaillée
            try:
                await wait_any([
                    f"#licence_contener_{user_id}",
                    f"#licence_mobile_list_{user_id}",
                    f"#link_licence_mobile_{user_id}",
                    f"#link_contract_mobile_{user_id}",
                ], timeout=15000)
            except PlaywrightTimeoutError:
                if page.is_closed():
                    print("⚠️ La page a été fermée pendant l’ouverture des détails.")
                    self.processed_clients.append({
                        "client_name": client_name,
                        "phone_number": target_number,
                        "rio_code": "Page fermée lors de l'ouverture des détails",
                        "status": "rio_empty",
                    })
                    return
                # petit réessai
                print("↻ Réessai ouverture des détails…")
                try:
                    await details_control.click()
                except Exception:
                    await page.evaluate("(el)=>el.click()", details_control)
                await wait_any([
                    f"#licence_contener_{user_id}",
                    f"#licence_mobile_list_{user_id}",
                    f"#link_licence_mobile_{user_id}",
                    f"#link_contract_mobile_{user_id}",
                ], timeout=8000)

            print(f"✅ Détails ouverts pour {target_number}")

            # Extraction robuste via la fonction dédiée
            rio_code = await self.open_mobile_and_get_rio(page, user_id, target_number)

            if rio_code and rio_code.strip():
                rio = rio_code.strip()
                print(f"✅ RIO extrait: {rio} pour {target_number}")
                self.save_rio_to_excel(client_name, target_number, rio)
                self.processed_clients.append({
                    "client_name": client_name,
                    "phone_number": target_number,
                    "rio_code": rio,
                    "status": "rio_extracted",
                })
            else:
                print(f"❌ RIO introuvable pour {target_number}")
                self.processed_clients.append({
                    "client_name": client_name,
                    "phone_number": target_number,
                    "rio_code": "RIO field empty or inaccessible",
                    "status": "rio_empty",
                })

        except Exception as e:
            print(f"❌ Erreur RIO ({target_number}): {e}")

    async def open_mobile_and_get_rio(self, page: Page, user_id: str, targetPhone: Optional[str] = None) -> Optional[str]:
        """Ouvre 'Lignes mobiles' -> clique la ligne (par numéro si dispo) -> Onglet Contrat -> lit le RIO."""
        async def wait_any(selectors: List[str], timeout: int = 15000):
            js = """
            sels => {
            for (const sel of sels) {
                const el = document.querySelector(sel);
                if (!el) continue;
                const cs = getComputedStyle(el);
                if (!cs) continue;
                if (el.offsetParent !== null || cs.display !== 'none' || cs.visibility !== 'hidden') return true;
            }
            return false;
            }
            """
            return await page.wait_for_function(js, arg=selectors, timeout=timeout)

        try:
            # 1) Dérouler "Lignes mobiles"
            list_link = page.locator(f"#link_licence_mobile_list_{user_id}")
            await list_link.scroll_into_view_if_needed()

            try:
                expanded_attr = await list_link.get_attribute("aria-expanded")
                expanded = (expanded_attr or "").lower() == "true"
            except Exception:
                expanded = False

            if not expanded:
                try:
                    await list_link.click()
                except Exception:
                    try:
                        await page.evaluate("(el)=>el.click()", await list_link.element_handle())
                    except Exception:
                        print("🔧 Fallback LicenceMenuActivate (liste mobiles)")
                        await page.evaluate(
                            """(args) => {
                                const el = document.getElementById(args.id);
                                if (!el) return;
                                if (typeof window.LicenceMenuActivate === 'function') {
                                    window.LicenceMenuActivate(el, args.uid);
                                } else {
                                    el.dispatchEvent(new MouseEvent('click', {bubbles: true}));
                                }
                            }""",
                            {"id": f"link_licence_mobile_list_{user_id}", "uid": user_id},
                        )

            # attendre ouverture du collapse liste (avec réessai)
            try:
                await page.wait_for_function(
                    """sel => {
                        const el = document.querySelector(sel);
                        if (!el) return false;
                        const cs = getComputedStyle(el);
                        return el.classList.contains('in') || el.classList.contains('show') || cs.display !== 'none';
                    }""",
                    arg=f"#licence_mobile_list_{user_id}",
                    timeout=12000,
                )
            except PlaywrightTimeoutError:
                try:
                    await list_link.click()
                except Exception:
                    with suppress(Exception):
                        await page.evaluate("(el)=>el.click()", await list_link.element_handle())
                await page.wait_for_timeout(350)
                await page.wait_for_function(
                    """sel => {
                        const el = document.querySelector(sel);
                        if (!el) return false;
                        const cs = getComputedStyle(el);
                        return el.classList.contains('in') || el.classList.contains('show') || cs.display !== 'none';
                    }""",
                    arg=f"#licence_mobile_list_{user_id}",
                    timeout=8000,
                )
            print(f"📂 Lignes mobiles ouvertes (user {user_id})")

            # 2) Cliquer la ligne mobile
            if targetPhone:
                norm_digits = self._normalize_phone(targetPhone)
                pat = self._spaced_pattern(norm_digits)  # ex: 0\s*6\s*07…
                mobile_line = page.locator(f"#licence_mobile_list_{user_id} a.menu-mobile-line").filter(has_text=pat)
                if await mobile_line.count() == 0:
                    print(f"⚠️ Ligne avec {norm_digits} introuvable, fallback par ID")
                    mobile_line = page.locator(f"#licence_mobile_list_{user_id} a#link_licence_mobile_{user_id}")
            else:
                mobile_line = page.locator(f"#licence_mobile_list_{user_id} a#link_licence_mobile_{user_id}")

            await mobile_line.first.scroll_into_view_if_needed()
            try:
                await mobile_line.first.click()
            except Exception:
                try:
                    await page.evaluate("(el)=>el.click()", await mobile_line.first.element_handle())
                except Exception:
                    print("🔧 Fallback LicenceMenuActivate (ligne mobile)")
                    await page.evaluate(
                        """(args) => {
                            const el = document.getElementById(args.id);
                            if (!el) return;
                            if (typeof window.LicenceMenuActivate === 'function') {
                                window.LicenceMenuActivate(el, args.uid);
                            } else {
                                el.dispatchEvent(new MouseEvent('click', {bubbles: true}));
                            }
                        }""",
                        {"id": f"link_licence_mobile_{user_id}", "uid": user_id},
                    )

            # attendre ouverture du panneau de la ligne (avec réessai)
            try:
                await page.wait_for_function(
                    """sel => {
                        const el = document.querySelector(sel);
                        if (!el) return false;
                        const cs = getComputedStyle(el);
                        return el.classList.contains('in') || el.classList.contains('show') || cs.display !== 'none';
                    }""",
                    arg=f"#licence_mobile_{user_id}",
                    timeout=12000,
                )
            except PlaywrightTimeoutError:
                # petit réessai sur le clic de la ligne
                try:
                    await mobile_line.first.click()
                except Exception:
                    with suppress(Exception):
                        await page.evaluate("(el)=>el.click()", await mobile_line.first.element_handle())
                await page.wait_for_timeout(350)
                await page.wait_for_function(
                    """sel => {
                        const el = document.querySelector(sel);
                        if (!el) return false;
                        const cs = getComputedStyle(el);
                        return el.classList.contains('in') || el.classList.contains('show') || cs.display !== 'none';
                    }""",
                    arg=f"#licence_mobile_{user_id}",
                    timeout=8000,
                )
            print(f"📄 Panneau ligne mobile ouvert (user {user_id})")

            # 3) Onglet "Contrat"
            mobile_panel = page.locator(f"#licence_mobile_{user_id}")
            contract_tab = page.locator(f"#link_contract_mobile_{user_id}")
            if await contract_tab.count() == 0:
                contract_tab = mobile_panel.locator('a[role="tab"]').filter(has_text="Contrat")

            await contract_tab.first.scroll_into_view_if_needed()
            try:
                await contract_tab.first.click()
            except Exception:
                await page.evaluate("(el)=>el.click()", await contract_tab.first.element_handle())

            # petit délai pour le contenu
            await page.wait_for_timeout(400)

            # === NOUVEAU : lecture RIO (garde TOUT le flow précédent inchangé) ===
            rio = await self._read_rio_from_scope(page, mobile_panel, user_id)
            return rio

        except Exception as e:
            print(f"❌ Erreur open_mobile_and_get_rio: {e}")
            return None

    # Orchestration
    # =========================
    async def process_all_clients(self, page: Page, clients: List[str]):
        total = len(clients)
        print(f"\n🚀 Lancement: {total} client(s)")
        for i, client in enumerate(clients, 1):
            print(f"\n[{i}/{total}] {client}")
            ok = await self.search_and_click_client(page, client)
            if ok:
                print(f"✅ Terminé: {client}")
            else:
                print(f"❌ Échec: {client}")
            await page.wait_for_timeout(500)
        self.print_processing_summary()

    def print_processing_summary(self):
        client_events = [c for c in self.processed_clients if "phone_number" not in c]
        rio_events = [c for c in self.processed_clients if "phone_number" in c]

        total_clients = len(client_events)
        successful = sum(1 for c in client_events if c["status"] == "success")
        not_found = sum(1 for c in client_events if c["status"] == "not_found")
        errors = sum(1 for c in client_events if str(c["status"]).startswith("error"))

        rio_extracted = sum(1 for r in rio_events if r["status"] == "rio_extracted")
        rio_missing = sum(1 for r in rio_events if r["status"] != "rio_extracted")

        print("\n📈 RÉCAP:")
        print(f"Clients total: {total_clients}")
        print(f"✅ Clients ouverts: {successful}")
        print(f"❌ Clients introuvables: {not_found}")
        print(f"⚠️ Erreurs: {errors}")
        print(f"📱 RIO extraits: {rio_extracted}")
        print(f"🚫 RIO manquants: {rio_missing}")

        if self.processed_clients:
            print("\n📋 Détails:")
            for item in self.processed_clients:
                if "phone_number" not in item:
                    icon = "✅" if item["status"] == "success" else "❌"
                    print(f"{icon} {item['client_name']} - {item['status']} {('[' + item.get('url','') + ']') if item.get('url') else ''}")
                else:
                    print(f"   📞 {item['phone_number']}: {item.get('rio_code','')} ({item['status']})")

    async def search_client(self, page: Page, search_term: str):
        print(f"Recherche manuelle: {search_term}")
        await page.fill("input#customer_list_search", "")
        await page.type("input#customer_list_search", search_term, delay=80)
        await page.wait_for_timeout(1200)
        print("✔️ Résultats affichés dans le tableau.")

    async def run_automation(self):
        async with async_playwright() as pw:
            browser, context, page = await self.setup_browser(pw)
            try:
                await self.login(context, page)

                if await self.navigate_to_client_page(page):
                    clients = self.read_clients_from_excel()
                    if clients:
                        await self.process_all_clients(page, clients)
                    else:
                        print("\n⚠️ Pas de clients en Excel. Mode recherche manuelle : tape un terme dans la console du script si tu ajoutes cette partie, ou utilise directement le champ dans le navigateur.")
                else:
                    print("❌ Impossible d'atteindre la page client.")
            except Exception as e:
                print(f"💥 Erreur run_automation: {e}")
            finally:
                await browser.close()


# =========================
# Entrée principale
# =========================
async def main():
    try:
        from dotenv import load_dotenv  # type: ignore
        load_dotenv()
    except Exception:
        print("ℹ️ Conseil: installe `python-dotenv` et crée un fichier .env pour charger UNYC_USERNAME/UNYC_PASSWORD")
    print("🚀 UNYC Atlas Automation (Playwright)")
    print("Ctrl+C pour arrêter.")
    automation = UnycAutomation()
    await automation.run_automation()




if __name__ == "__main__":
    asyncio.run(main())
