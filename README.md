#  Excel Search App / Lokalna Wyszukiwarka Excel

[![en](https://img.shields.io/badge/lang-English-red.svg)](#-english-version)
[![pl](https://img.shields.io/badge/lang-Polski-white.svg)](#-polska-wersja)

---

## 🇬🇧 English Version

### About the Project
**Excel Search App** (Wyszukiwarka Excel) is a lightweight, fast Windows Desktop application written in C# (Windows Forms, .NET Core 3.1). It allows users to quickly load Excel spreadsheets (`.xls`, `.xlsx`) and perform real-time, fuzzy searches across all columns simultaneously. 

It is especially useful for quickly finding people, items, or specific records in large datasets without needing to open Microsoft Excel.

###  Features
* **Real-time Filtering:** The data grid updates instantly as you type.
* **Smart Search:** The search engine automatically ignores letter casing, punctuation, and Polish diacritical marks (e.g., typing "zolc" will match "żółć").
* **Safe Reading:** Opens files in "ReadWrite" mode, meaning you can search through an Excel file even if it is currently opened in Microsoft Excel.
* **No Excel Required:** Powered by `ExcelDataReader`, it does not require Microsoft Office to be installed on the machine.

###  How to Use
1. **Launch the App:** Run the compiled `.exe` file.
2. **Load a File:** Click the blue **"Załaduj plik Excel"** (Load Excel file) button in the top left corner.
3. **Select a Spreadsheet:** Choose any `.xlsx` or `.xls` file from your computer.
4. **Search:** Once the data is loaded, type your query into the search bar at the top (` Wpisz dowolne dane...`). The results will filter automatically.

###  Tech Stack & Requirements
* **Framework:** .NET Core 3.1 (Windows Desktop)
* **UI:** Windows Forms
* **Dependencies:** [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) & ExcelDataReader.DataSet
* **To run the app:** You need the [.NET Core 3.1 Desktop Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/3.1) installed.

###  How to Build 
1. Clone the repository.
2. Open the `.csproj` file in Visual Studio.
3. Restore NuGet packages (it should happen automatically).
4. Build and Run.

---

## 🇵🇱 Polska Wersja

### O projekcie
**Wyszukiwarka Excel** to lekka i szybka aplikacja okienkowa (Windows Desktop) napisana w języku C# (Windows Forms, .NET Core 3.1). Pozwala na błyskawiczne wczytywanie arkuszy kalkulacyjnych (`.xls`, `.xlsx`) i wyszukiwanie danych w czasie rzeczywistym we wszystkich kolumnach jednocześnie.

Aplikacja jest idealna do szybkiego przeszukiwania list osób, przedmiotów czy rekordów bez konieczności uruchamiania programu Microsoft Excel.

###  Główne funkcje
* **Filtrowanie na żywo:** Tabela z wynikami aktualizuje się natychmiast podczas wpisywania tekstu.
* **Inteligentne wyszukiwanie:** Wyszukiwarka ignoruje wielkość liter, interpunkcję oraz **polskie znaki** (np. wpisanie "zolc" znajdzie "żółć").
* **Bezpieczny odczyt:** Pliki otwierane są w trybie bezpiecznym, co pozwala na przeszukiwanie pliku nawet wtedy, gdy jest on jednocześnie otwarty w programie Excel.
* **Brak wymogu MS Office:** Dzięki bibliotece `ExcelDataReader`, aplikacja nie wymaga zainstalowanego pakietu Microsoft Office na komputerze.

###  Jak używać
1. **Uruchom aplikację:** Włącz skompilowany plik `.exe`.
2. **Wczytaj plik:** Kliknij niebieski przycisk **"Załaduj plik Excel"** w lewym górnym rogu.
3. **Wybierz arkusz:** Wybierz dowolny plik `.xlsx` lub `.xls` z dysku.
4. **Wyszukaj:** Po załadowaniu danych, zacznij wpisywać szukaną frazę w pasku wyszukiwania (` Wpisz dowolne dane, aby wyszukać...`). Wyniki będą filtrować się automatycznie.

###  Technologie i Wymagania
* **Środowisko:** .NET Core 3.1 (Windows Desktop)
* **Interfejs:** Windows Forms
* **Zależności:** [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) oraz ExcelDataReader.DataSet
* **Do uruchomienia:** Wymagane jest środowisko [.NET Core 3.1 Desktop Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/3.1).

###  Jak skompilować
1. Sklonuj repozytorium.
2. Otwórz plik projektu (`.csproj` / `.sln`) w programie Visual Studio.
3. Przywróć pakiety NuGet (powinno nastąpić automatycznie).
4. Skompiluj i uruchom aplikację.
