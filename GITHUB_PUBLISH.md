# Как выложить проект на GitHub

Репозиторий уже инициализирован и сделан первый коммит. Осталось создать репозиторий на GitHub и отправить код.

## Шаг 1. Создайте репозиторий на GitHub

1. Зайдите на [github.com](https://github.com) и войдите в аккаунт.
2. Нажмите **«+»** → **«New repository»**.
3. Укажите имя, например: **KP_Generator** или **KP-Final**.
4. Описание (по желанию): «Генератор коммерческих предложений (Excel/PDF → Excel или Word)».
5. **Не** ставьте галочки «Add a README», «Add .gitignore» — они уже есть в проекте.
6. Нажмите **«Create repository»**.

## Шаг 2. Привяжите удалённый репозиторий и отправьте код

В терминале перейдите в папку проекта и выполните (подставьте **свой логин** и **имя репозитория**):

```bash
cd "e:\KP_Generator_Final_Project\KP_Final"

git remote add origin https://github.com/ВАШ_ЛОГИН/ИМЯ_РЕПОЗИТОРИЯ.git
git branch -M main
git push -u origin main
```

Пример, если логин `ivanov`, репозиторий `KP_Generator`:

```bash
git remote add origin https://github.com/ivanov/KP_Generator.git
git branch -M main
git push -u origin main
```

При первом `git push` браузер или Git могут запросить вход в GitHub — войдите и повторите команду при необходимости.

## Готово

После успешного `git push` ссылка для проверки будет такой:

**https://github.com/ВАШ_ЛОГИН/ИМЯ_РЕПОЗИТОРИЯ**

Её можно отправить для проверки.
