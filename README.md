# PS - L4

Aplikacja do filtrowania pliku Excel zawierających dane pracowników i zwolnienia lekarskie (L4). Program pozwala na wybór odpowiednich arkuszy z obu plików i generuje raport zawierający tylko wspólne rekordy, bazując na numerach PESEL.

## Funkcje

- Łączenie danych z dwóch plików Excel (XLSX/XLS)
- Filtrowanie danych na podstawie numerów PESEL
- Generowanie raportu w formacie Excel
- Logowanie operacji

## Instalacja

1. Sklonuj repozytorium
2. Przejdź do katalogu projektu
3. Uruchom komendę:
```bash
cargo build --release
```

## Użycie

1. Uruchom aplikację
2. Wybierz plik z listą pracowników
3. Wybierz odpowiedni arkusz z listą pracowników
4. Wybierz plik z danymi L4
5. Wybierz odpowiedni arkusz z danymi L4
6. (Opcjonalnie) Zmień nazwę pliku wynikowego
7. Kliknij "Uruchom"

## Format danych wejściowych

### Plik z listą pracowników
Powinien zawierać kolumny:
- Nazwisko
- Imię
- PESEL

### Plik z L4
Powinien zawierać kolumny:
- Ubezpieczony (w formacie: "Nazwisko Imię PESEL")
- Data od
- Data do
- Na opiekę
- Pobyt w szpitalu
- Status zaświadczenia

## Format danych wyjściowych

Plik wynikowy zawiera następujące kolumny:
- Nazwisko
- Imię
- PESEL
- Data od
- Data do
- Na opiekę
- Pobyt w szpitalu
- Status zaświadczenia

## Autor

Oleksii Sliepov 