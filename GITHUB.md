# Как залить на GitHub

## Вариант 1 — через сайт (проще)

1. Зайди на https://github.com/new и создай новый репозиторий,
   например `vyrabotka` (без README — оставь пустым).
2. На странице репозитория нажми **Add file → Upload files**.
3. Перетащи файлы: `app.py`, `processor.py`, `requirements.txt`,
   `README.md`, `.gitignore`, `build_exe.bat`.
4. Нажми **Commit changes**. Готово.

## Вариант 2 — через командную строку (git)

Открой терминал в папке с этими файлами и выполни:

```bash
git init
git add app.py processor.py requirements.txt README.md .gitignore build_exe.bat
git commit -m "Программа: выработка по сотрудникам"
git branch -M main

# подставь свой логин и название репозитория:
git remote add origin https://github.com/ТВОЙ_ЛОГИН/vyrabotka.git
git push -u origin main
```

При первом пуше GitHub попросит авторизацию. Введи логин и
**Personal Access Token** вместо пароля
(создаётся в Settings → Developer settings → Personal access tokens).

## Обновление в будущем

```bash
git add .
git commit -m "что изменил"
git push
```

## Автосборка .exe на GitHub

В репозитории есть файл `.github/workflows/build.yml`. После загрузки на GitHub:

1. Открой вкладку **Actions** в репозитории.
2. Запусти workflow **Build EXE** (или он запустится сам после `git push`).
3. Когда сборка закончится (зелёная галочка) — внизу страницы запуска,
   в разделе **Artifacts**, скачай **Vyrabotka-exe**.

Внутри будет готовый `Vyrabotka.exe` — Windows соберёт его за тебя, ставить
ничего не нужно.
