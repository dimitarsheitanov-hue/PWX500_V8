CAM Agent V5.1.38 BUILD FIX4 - nested batch browser regression fix

Причина: Етап 2 използва вложена папка ROOT\Общи_пакети\DATE-RANGE_TIMESTAMP. Старият генеричен път се опитваше да намери вложения folder като TreeItem в левия панел. Maestro WPF виртуализира дървото и може да показва папката само в централния списък.

Поправка (само за общи пакети): отваря по една директория от централния списък, проверява актуалния breadcrumb след всяко двойно щракване. Това се прилага в Add programs, Choose the file (PowerSet) и Save as. Не се натиска Confirm при неуспешна проверка.

Основният дневен режим, старият Desktop\FOUND\dd.MM navigator, конверторът, CSV/HOP/PGMX workers, времената, правилата за изключения и самите етапи са непроменени.

GitHub: качи ЦЯЛОТО съдържание в ново repository, включително .github/workflows/build-exe.yml; Actions -> Build CAM Agent V5.1.38 EXE -> Run workflow; изтегли и разархивирай пълния COMPILED_PACKAGE.zip. Стартирай EXE от пълната папка.

При тест: не повтаряй Етап 1. Общият batch manifest остава под ROOT\Общи_пакети\... . Пусни само 2. Общ MIXX + PowerSet + XLS. При неуспех запази POWERSET_AUTOMATION_LOG.txt и PWX500_AutoMix.log; не изтривай manifest / PGMX.

Тестове: PowerShell parser и чисти path/safety тестове в GitHub; реален Maestro/PowerSet Windows UI test се прави при оператора.
