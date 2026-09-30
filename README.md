# Rejestr czasu pracy — instrukcja wdrożenia od A do Z

Aplikacja przeglądarkowa do rejestrowania czasu pracy zespołu (do ok. 10 osób).
Dane trafiają do **pliku Excel na OneDrive**, który można w każdej chwili otworzyć,
przefiltrować, podsumować albo wydrukować w zwykłym Excelu.

**Pola wpisu:** Imię i nazwisko · Data · ID projektu (lista ze słownika) · Opis · Liczba godzin

---

## 0. Co jest w paczce

| Plik | Do czego służy |
|---|---|
| `index.html` | Aplikacja (wygląd). |
| `app.js` | Logika: logowanie, wpisy, filtry, podsumowania. |
| `xlsx-store.js` | Odczyt i zapis tabel w pliku .xlsx (zmienia tylko wiersze z wpisami, formatowanie zostaje). |
| `config.js` | **Twoja konfiguracja**, czyli jedyny plik, który edytujesz. |
| `lib/` | Biblioteki: logowanie Microsoft (MSAL) i obsługa ZIP (JSZip). Działają bez internetowego CDN. |
| `CzasPracy.xlsx` | Szablon bazy danych: arkusze **Wpisy**, **Projekty**, **Podsumowanie**. |

### Jak to działa

```
Przeglądarka ──logowanie kontem Microsoft──► Microsoft Entra ID
     │
     └──► Microsoft Graph ──► OneDrive: CzasPracy.xlsx
            (pobierz plik → dopisz wiersz → zapisz z kontrolą wersji)
```

Aplikacja nie ma własnego serwera ani bazy. Dostęp do danych ma tylko osoba, której
udostępniono plik na OneDrive. Jeśli dwie osoby zapiszą coś w tej samej chwili, aplikacja wykryje
konflikt (nagłówek `If-Match`), pobierze świeżą wersję i ponowi zapis, więc żaden wpis nie zginie.

---

## 1. Wypróbuj aplikację (5 minut, bez konfiguracji)

Tryb demo działa bez logowania, a dane zapisuje tylko w Twojej przeglądarce.

1. Rozpakuj folder `czas-pracy`.
2. Uruchom w nim prosty serwer WWW. Otwarcie `index.html` dwuklikiem nie zadziała, bo przeglądarka blokuje wtedy wczytanie pliku Excel.
   - Windows: w folderze wpisz w pasku adresu Eksploratora `cmd`, a następnie:
     `python -m http.server 8000`
     (gdy brak Pythona: zainstaluj go z Microsoft Store albo użyj rozszerzenia *Live Server* w VS Code).
3. Wejdź na **http://localhost:8000/** i kliknij **Uruchom tryb demo**.
4. Przyciskiem **Pobierz plik Excel** na żółtym pasku możesz zobaczyć, jak wpisy wyglądają w Excelu.

---

## 2. Przygotuj plik Excel na OneDrive

1. Wgraj `CzasPracy.xlsx` do swojego OneDrive (np. do folderu `Rejestr czasu pracy`).
2. Otwórz arkusz **Projekty** i wpisz swoje projekty w tabeli:
   - **ID projektu**: krótki kod widoczny na liście w aplikacji, np. `P-2026-014`.
   - **Nazwa projektu**: opis wyświetlany obok kodu.
   - **Aktywny**: `T` pokazuje projekt na liście wyboru, `N` go ukrywa, ale stare wpisy zostają.
   Przykładowe projekty (P-001, P-002…) możesz usunąć albo nadpisać.
3. **Udostępnij plik** zespołowi (maks. ok. 10 osób):
   - kliknij plik prawym przyciskiem → **Udostępnij**;
   - wpisz adresy e-mail kont Microsoft współpracowników (outlook.com, hotmail.com, live.com albo konta firmowe);
   - ustaw uprawnienie **Może edytować**;
   - wyślij zaproszenie.
4. Skopiuj **link do pliku**: prawy przycisk → **Kopiuj link**, z uprawnieniem *Może edytować*.
   Link wygląda np. tak: `https://1drv.ms/x/c/…` albo `https://…-my.sharepoint.com/:x:/g/personal/…`.
   Wkleisz go w kroku 5 jako `shareUrl`.

> Bezpieczeństwo: najlepiej ustaw link „Określone osoby”, a nie „Każdy, kto ma link”.
> Aplikacja i tak wymaga zalogowania, ale zawężony link to dodatkowa ochrona.

---

## 3. Opublikuj aplikację w internecie (bezpłatnie: GitHub Pages)

Aplikacja to kilka plików statycznych i **nie zawiera żadnych haseł**, więc może być publiczna.
Identyfikator aplikacji (`clientId`) nie jest tajny. Dane chroni logowanie i uprawnienia do pliku.

1. Załóż konto na https://github.com (bezpłatne).
2. **New repository** → nazwa np. `czas-pracy` → *Public* → **Create repository**.
3. **Add file → Upload files** → przeciągnij **zawartość** folderu `czas-pracy`
   (`index.html`, `app.js`, `xlsx-store.js`, `config.js`, folder `lib`) → **Commit changes**.
   Pliku `CzasPracy.xlsx` nie musisz wgrywać, bo jest potrzebny tylko w trybie demo.
4. **Settings → Pages** → *Source*: **Deploy from a branch** → *Branch*: `main` / `(root)` → **Save**.
5. Po 1–2 minutach aplikacja będzie pod adresem:
   **`https://<twoja-nazwa>.github.io/czas-pracy/`**. Zapisz go, to Twój **adres przekierowania** do kroku 4.

Inne możliwości hostingu: Azure Static Web Apps (plan Free), Netlify, serwer firmowy.
Warunek jest jeden: strona musi działać przez **HTTPS**.

---

## 4. Zarejestruj aplikację w Microsoft Entra ID

Ten krok pozwala aplikacji logować użytkowników kontem Microsoft i zapisywać do pliku na OneDrive.

### 4a. Potrzebujesz katalogu (tenanta)

- **Konto firmowe / szkolne Microsoft 365**: katalog już masz. Rejestrację może zrobić
  Ty (jeśli firma na to pozwala) albo dział IT.
- **Konto prywatne (outlook.com, hotmail.com)**: Microsoft nie pozwala już rejestrować
  aplikacji bez katalogu. Załóż **bezpłatne konto Azure** (https://azure.microsoft.com/free),
  logując się tym samym kontem. Powstanie wtedy „Default Directory”. Rejestracja aplikacji
  jest bezpłatna i nie zużywa środków.

### 4b. Rejestracja

1. Wejdź na https://entra.microsoft.com → **Identity → Applications → App registrations → New registration**.
2. **Name**: `Rejestr czasu pracy`.
3. **Supported account types**: wybierz
   **„Accounts in any organizational directory and personal Microsoft accounts”**.
   Dzięki temu zalogują się zarówno konta prywatne, jak i firmowe.
4. **Redirect URI**: platforma **Single-page application (SPA)**, adres z kroku 3, np.
   `https://twoja-nazwa.github.io/czas-pracy/`
   Adres musi się zgadzać **co do znaku**, łącznie z ukośnikiem na końcu.
5. **Register**.
6. Na stronie *Overview* skopiuj **Application (client) ID**, czyli ciąg w stylu `1a2b3c4d-…`.
7. **Authentication** → w sekcji *Single-page application* dodaj drugi adres do testów lokalnych:
   `http://localhost:8000/` → **Save**.
8. **API permissions → Add a permission → Microsoft Graph → Delegated permissions** → zaznacz
   `Files.ReadWrite.All` (oraz `User.Read`, zwykle jest już dodane) → **Add permissions**.
   W katalogu firmowym administrator może od razu kliknąć **Grant admin consent**.

> Dlaczego `Files.ReadWrite.All`? Pozwala otworzyć plik, który **ktoś inny** udostępnił.
> Węższe uprawnienie `Files.ReadWrite` obejmuje tylko własne pliki.
> Aplikacja korzysta wyłącznie z pliku wskazanego w `config.js`, a działa zawsze w imieniu zalogowanej osoby i w granicach jej uprawnień.

---

## 5. Uzupełnij `config.js`

Otwórz `config.js` w Notatniku i wpisz:

```js
window.CZAS_PRACY_CONFIG = {
  clientId: "1a2b3c4d-....",                       // z kroku 4b.6
  shareUrl: "https://1drv.ms/x/c/....",            // z kroku 2.4
  admins: ["p.piechota@outlook.com"],              // kto może poprawiać wpisy innych
  appName: "Rejestr czasu pracy",
  maxHoursPerDayWarning: 12,
};
```

Wgraj poprawiony `config.js` do repozytorium GitHub (**Add file → Upload files**, nadpisz plik).
Po ok. minucie zmiany będą widoczne na stronie.

---

## 6. Pierwsze uruchomienie

1. Wejdź na adres aplikacji → **Zaloguj się kontem Microsoft**.
2. Przy pierwszym logowaniu Microsoft zapyta o zgodę na dostęp do plików. Zaakceptuj.
3. Aplikacja wczyta plik i pokaże formularz.
4. Wyślij zespołowi link do aplikacji. Każda osoba loguje się **tym kontem, któremu udostępniono plik**.

Pierwsze logowanie warto przetestować z jednym współpracownikiem, zanim roześlesz link do wszystkich.

---

## 7. Codzienna obsługa

**Dodawanie wpisu:** data (domyślnie dziś), godziny (np. `7,5` lub `7.5`), imię i nazwisko
(podpowiada się z logowania i jest zapamiętywane), projekt z listy, opis → **Zapisz wpis**.

**Lista wpisów:** filtry miesiąca (strzałki ‹ ›), osoby i projektu, a u góry suma godzin.
Przy każdym wpisie są przyciski:
- ⧉ **Kopiuj**: tworzy nowy wpis z tym samym projektem i opisem na dziś, co przydaje się przy powtarzalnej pracy;
- ✎ **Edytuj** i 🗑 **Usuń**: dostępne dla autora wpisu i administratorów.

**Podsumowanie:** godziny wg projektu i wg osoby dla wybranych filtrów.

**Kontrole:** godziny muszą być w zakresie 0–24, a suma godzin jednej osoby w jednym dniu nie może
przekroczyć 24. Powyżej 12 h aplikacja poprosi o potwierdzenie (próg ustawiasz w `config.js`).

**Excel:** przycisk **Excel ↗** otwiera plik. Arkusz **Podsumowanie** liczy godziny wg projektu
(i wg osób, które wpiszesz w kolumnie E) dla okresu podanego w komórkach B3–B4.

**Nowy projekt:** dopisz wiersz w arkuszu **Projekty**. Pojawi się w aplikacji po odświeżeniu (↻).

---

## 8. Zasady, o których trzeba pamiętać

1. **Nie zmieniaj nazw tabel** (`tblWpisy`, `tblProjekty`) ani nagłówków kolumn.
   Kolejność kolumn możesz zmieniać. Możesz też dodać własne kolumny w tabeli Wpisy, a aplikacja je zachowa.
2. **Nie wpisuj niczego pod tabelą Wpisy** w kolumnach A–H, bo aplikacja przy zapisie przebudowuje ten obszar.
   Notatki umieszczaj w innych arkuszach albo w kolumnach na prawo od tabeli.
3. **Excel na komputerze (desktop) może zablokować plik**, dopóki jest w nim otwarty do edycji. Aplikacja pokaże
   wtedy komunikat „Plik jest zablokowany”. Do przeglądania używaj Excela w przeglądarce albo
   zamykaj plik po pracy.
4. Poprawki wprowadzone ręcznie w Excelu są widoczne w aplikacji po odświeżeniu.
   Wiersze dopisane ręcznie dostaną identyfikator automatycznie przy najbliższym zapisie z aplikacji.
5. **Kopie zapasowe:** OneDrive przechowuje historię wersji pliku
   (prawy przycisk → **Historia wersji**), więc zawsze można przywrócić wcześniejszy stan.
6. Ograniczenie uprawnień do edycji własnych wpisów działa w aplikacji. Osoba, która ma
   prawo edycji pliku, może go też zmienić bezpośrednio w Excelu. Przy 10 zaufanych osobach to zwykle wystarcza.

---

## 9. Rozwiązywanie problemów

| Komunikat / objaw | Przyczyna i rozwiązanie |
|---|---|
| `AADSTS50011` … redirect URI mismatch | Adres w Entra ID różni się od adresu strony. Skopiuj adres z paska przeglądarki (bez `?…` i `#…`, z ukośnikiem na końcu) do *Authentication → Single-page application*. |
| `AADSTS700016` … application not found | Błędny `clientId` w `config.js` albo aplikacja nie obsługuje kont prywatnych (krok 4b.3). |
| `AADSTS50020` / „konto nie istnieje w dzierżawie” | W kroku 4b.3 wybrano „tylko ten katalog”. Zmień w *Authentication → Supported account types* albo w *Manifest* (`signInAudience`: `AzureADandPersonalMicrosoftAccount`). |
| „Wymagana zgoda administratora” (konto firmowe) | Firma blokuje zgody użytkowników. Poproś IT o **Grant admin consent** dla aplikacji. |
| „Brak dostępu do pliku Excel (403/404)” | Plik nie jest udostępniony temu kontu albo link w `shareUrl` jest nieaktualny. Udostępnij plik ponownie (z edycją) i wygeneruj nowy link. |
| „Plik jest zablokowany” | Ktoś ma plik otwarty w Excelu na komputerze. Trzeba go zamknąć. |
| „W tabeli tblWpisy brakuje kolumn…” | Zmieniono nazwę kolumny lub tabeli w Excelu. Przywróć nazwy z szablonu. |
| Lista projektów jest pusta | W arkuszu Projekty brak wierszy z `Aktywny = T` albo wpisano je poza tabelą. Rozszerz tabelę (przeciągnij róg tabeli). |
| Po zmianie `config.js` nic się nie zmienia | Odśwież stronę z pominięciem pamięci podręcznej: `Ctrl+F5`. |

---

## 10. Możliwe rozszerzenia

- Blokada edycji zamkniętych miesięcy (np. po 5. dniu kolejnego miesiąca).
- Eksport miesięcznej karty czasu pracy do PDF dla każdej osoby.
- Przeniesienie pliku na SharePoint (Microsoft 365 dla firm), co pozwala zapisywać także przy otwartym pliku w Excelu.
