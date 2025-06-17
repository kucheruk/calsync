# Безопасность и работа с секретами

## 🔒 Важные принципы безопасности

### Никогда не коммитьте реальные секреты!

Этот проект содержит инфраструктуру для работы с Exchange Web Services, включая аутентификацию. **Никогда не добавляйте в git реальные учетные данные!**

## 📁 Файлы конфигурации

### ✅ Безопасные файлы (включены в git):
- `appsettings.json` - основная конфигурация без секретов
- `appsettings.Local.json.example` - пример локальной конфигурации
- Все тестовые файлы с мок-данными

### ❌ Файлы с секретами (исключены из git):
- `appsettings.Local.json` - ваша локальная конфигурация с реальными данными
- `appsettings.Development.json` - конфигурация для разработки
- Любые файлы с реальными паролями, токенами, URL

## 🛠️ Настройка локальной среды

1. **Скопируйте пример конфигурации:**
   ```bash
   cp calsync/appsettings.Local.json.example calsync/appsettings.Local.json
   ```

2. **Заполните реальные данные в `appsettings.Local.json`:**
   ```json
   {
       "Exchange": {
           "ServiceUrl": "https://your-exchange-server.com/EWS/Exchange.asmx",
           "Domain": "your-domain",
           "Username": "your-username",
           "Password": "your-password"
       }
   }
   ```

3. **Убедитесь, что файл исключен из git:**
   ```bash
   git status
   # appsettings.Local.json не должен появляться в списке изменений
   ```

## 🧪 Тестирование

### Мок-данные в тестах
Все тесты используют только мок-данные:
- **Домены:** `test.local`, `example.domain`
- **Пользователи:** `testuser`, `mockuser`
- **Серверы:** `localhost`, `exchange.example.com`
- **ID календарей:** `MOCK_CALENDAR_ID_*`

### Реальные интеграционные тесты
Тесты с префиксом `ExchangeIntegration*`:
- Автоматически пропускаются, если нет локальной конфигурации
- Используют реальные данные только локально
- Создают только тестовые события с префиксом `[TEST]`
- Автоматически очищают созданные данные

## 🔍 Проверка на секреты

Перед коммитом всегда проверяйте:

```bash
# Поиск потенциальных секретов
grep -r "password\|secret\|token\|key" . --exclude-dir=.git --exclude-dir=bin --exclude-dir=obj

# Проверка реальных доменов/серверов (замените на ваши реальные домены)
grep -r "your-domain\.com\|your-server\." . --exclude-dir=.git --exclude-dir=bin --exclude-dir=obj

# Проверка реальных пользователей (замените на ваши реальные имена)
grep -r "your-username\." . --exclude-dir=.git --exclude-dir=bin --exclude-dir=obj
```

## 🚨 Что делать, если секрет попал в git

1. **Немедленно смените пароль/токен**
2. **Удалите секрет из истории git:**
   ```bash
   git filter-branch --force --index-filter \
   'git rm --cached --ignore-unmatch path/to/file' \
   --prune-empty --tag-name-filter cat -- --all
   ```
3. **Force push (осторожно!):**
   ```bash
   git push origin --force --all
   ```

## ✅ Чек-лист перед коммитом

- [ ] Нет реальных паролей в коде
- [ ] Нет реальных URL серверов в тестах
- [ ] Нет реальных пользователей в тестах
- [ ] `appsettings.Local.json` не добавлен в git
- [ ] Все тесты используют мок-данные
- [ ] Документация не содержит реальных секретов

## 📞 Контакты

Если вы случайно закоммитили секреты или у вас есть вопросы по безопасности, немедленно свяжитесь с командой разработки. 