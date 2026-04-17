# Hledger-Excel Refactor - Test Checklist

## Pre-Test Setup
- [ ] Main.xlsm dosyası kapat (varsa açıksa)
- [ ] VBA kaynak dosyaları güncellenmiş: ✅ Config.bas, LogManager.bas, BankGetter refactored
- [ ] `C:\Budgeting\DATA\` klasörü mevcut ve yazılabilir

---

## Phase 1: Config System Test

**Test 1.1: Config Loading**
```
1. Main.xlsm'i Excel'de aç
2. Alt+F11 → VBA Editor
3. Immediate Window (Ctrl+G): Print Config.DATA_FOLDER
✅ Expected: C:\Budgeting\DATA\
```

**Test 1.2: All Config Constants**
```
Immediate Window'da:
- Print Config.HLEDGER_FILE_ADDR
- Print Config.TEMP_FILE_ADDR
- Print Config.PORTFOLIO_CSV_PATH
✅ Expected: All paths should resolve correctly
```

---

## Phase 2: Logging System Test

**Test 2.1: LogManager Initialization**
```
1. Excel başlat (Workbook_Open tetiklenecek)
2. Sheet'ler arasında "LOGS" sheet'i kontrol et
✅ Expected: LOGS sheet oluşturulmuş, header row var
   (Timestamp | Level | Message | Details)
```

**Test 2.2: Manual Logging Test**
```
Immediate Window:
Call LogManager.LogInfo("Test log message")
Call LogManager.LogDebug("Debug test")
Call LogManager.LogWarning("Warning test")
Call LogManager.LogError("Error test")

Sheet: LOGS'a git
✅ Expected: 4 satır loglanmış, renk-coded (blue/orange/red)
```

---

## Phase 3: CSV Import Test

**Test 3.1: ImporterBegin() Flow**
```
1. Bir test CSV hazırla (Date | Description | Amount)
2. Bank hesap sheet'ine CSV'yi yapıştır (header + 3 rows data)
3. Alt+Shift+Ctrl+I veya Code → Importer.ImporterBegin()
4. UI prompts'ları cevapla (Date/Description/Amount columns seç)
```

**Test 3.2: Rules Application & Logging**
```
Importer başlarken:
- LOGS sheet'te "CSV Import Process Started" görünsün mü?
✅ Expected: Import başladığında log entry olması
```

**Test 3.3: Rule Matching**
```
Eğer Rules sheet'te rule varsa:
- İşlem otomatik kategorize edilmeli
- LOGS'ta rule match info'su olmalı
```

---

## Phase 4: UTF-8 Validation Test

**Test 4.1: Run UTF-8 Round-Trip**
```
Immediate Window:
Call TestUTF8RoundTrip()

Bekle → "Data fetch completed" mesajı
```

**Test 4.2: Check Results**
```
1. UTF8Test sheet'ine git
2. 5 test case'ı kontrol et
✅ Expected: Column D'de "✓ PASS" olması
   (Column B: Original, Column C: Read Back, Column D: Match?)

Eğer FAIL varsa:
- LOGS sheet'te warning görülebilir
- Character encoding issue var demektir
```

---

## Phase 5: Hledger Export Test

**Test 5.1: CreateAllFilesAKATornado()**
```
1. MAIN_LEDGER sheet'te birkaç işlem olduğundan emin ol
2. Alt+Shift+Ctrl+X veya Code → MainModule.CreateAllFilesAKATornado()
3. Bekleme: Process hledger-ui açacak
```

**Test 5.2: Log Monitoring**
```
Main işlemi başlarken LOGS sheet'i watch et:
✅ Expected logs:
   - "Hledger File Generation Started"
   - "Account Transactions fetched" 
   - "Hledger File Generation Completed"
```

**Test 5.3: Output Files**
```
C:\Budgeting\DATA\ klasörüne git:
- Main.hledger dosyası oluşturulmuş mı?
- Türkçe karakterler koruyunmuş mu?
✅ Expected: Dosya var, UTF-8 encoded (açıp Türkçe görülmeli)
```

---

## Phase 6: BankGetter Automation Test

**Test 6.1: BankGetter_FetchAccounts()**
```
Eğer TEB online banking active'se:
1. Code → BankGetter_FetchAccounts(chrome)
   (veya MainModule.BankGetterTEB() bütün flow'ı çalıştırır)

2. Bank_Info sheet'ine bakılır
```

**Test 6.2: Logging Check**
```
LOGS sheet:
- "BankGetter: TEB Data Fetch Started" 
- "Fetching Account Transactions..."
✅ Expected: Each phase logged
```

---

## Error Handling Test

**Test 7.1: Force an Error**
```
1. C:\Budgeting\DATA\ klasörünü (geçici olarak) rename et
2. CreateAllFilesAKATornado() çalıştır
3. Error message oluşmalı
4. LOGS sheet'te ERROR entry olmalı
✅ Expected: Graceful error handling + logging
```

**Test 7.2: Restore**
```
Klasörü geri rename et
```

---

## Summary

| Test | Status | Notes |
|------|--------|-------|
| Config Loading | ? | Path constants OK mi? |
| LogManager Init | ? | LOGS sheet oluştuyor mu? |
| CSV Import + Rules | ? | Logging çalışıyor mu? |
| UTF-8 Encoding | ? | Turkish chars intact? |
| Hledger Export | ? | Files generated + encoded? |
| BankGetter Logging | ? | Optional, TEB banking aktif mi? |
| Error Handling | ? | Graceful failures? |

---

## Notes
- Logging'i check et: LOGS sheet'te her phase görülmeli
- UTF-8 test önemli: Türkçe karakterler kaybedilmemelidir
- BankGetter aktif kullanımdaysa, özel test gerekir

**Test sonuçlarını rapor ver!**
