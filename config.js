/*
 * KONFIGURACJA APLIKACJI — uzupełnij po wykonaniu kroków z pliku README.md
 */
window.CZAS_PRACY_CONFIG = {
  // Identyfikator aplikacji (Application (client) ID) z rejestracji w Microsoft Entra ID.
  // Pozostaw pusty, aby uruchomić tryb demonstracyjny (bez logowania, dane tylko w przeglądarce).
  clientId: "",

  // Link udostępniania do pliku CzasPracy.xlsx na OneDrive (z uprawnieniem do EDYCJI),
  // np. "https://1drv.ms/x/c/abc123/EXyz..." albo "https://firma-my.sharepoint.com/:x:/g/personal/...".
  shareUrl: "",

  // Adresy e-mail administratorów — mogą edytować i usuwać wpisy wszystkich osób.
  // Pozostałe osoby edytują i usuwają tylko wpisy, które same dodały.
  admins: [
    // "p.piechota@outlook.com",
  ],

  // Nazwa wyświetlana w nagłówku aplikacji.
  appName: "Rejestr czasu pracy",

  // Maksymalna liczba godzin jednej osoby w jednym dniu (powyżej — ostrzeżenie).
  maxHoursPerDayWarning: 12,
};
