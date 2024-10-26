Для установки Docker и Docker Compose на Ubuntu, следуйте этим шагам:

### Шаг 1: Обновление системы

Перед установкой обновим список пакетов и установим обновления:

```bash
sudo apt update
sudo apt upgrade
```

### Шаг 2: Установка Docker

1. **Удалите старые версии Docker, если они есть**:

   ```bash
   sudo apt remove docker docker-engine docker.io containerd runc
   ```

2. **Установите необходимые пакеты** для работы с репозиториями через HTTPS:

   ```bash
   sudo apt install apt-transport-https ca-certificates curl software-properties-common
   ```

3. **Добавьте официальный GPG-ключ Docker**:

   ```bash
   curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /usr/share/keyrings/docker-archive-keyring.gpg
   ```

4. **Добавьте репозиторий Docker**:

   ```bash
   echo "deb [arch=$(dpkg --print-architecture) signed-by=/usr/share/keyrings/docker-archive-keyring.gpg] https://download.docker.com/linux/ubuntu $(lsb_release -cs) stable" | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
   ```

5. **Установите Docker**:

   ```bash
   sudo apt update
   sudo apt install docker-ce docker-ce-cli containerd.io
   ```

6. **Проверьте установку**:

   ```bash
   sudo systemctl status docker
   ```

   Docker должен быть активен (Active: active (running)).

7. **Добавьте текущего пользователя в группу Docker**, чтобы можно было использовать Docker без `sudo`:

   ```bash
   sudo usermod -aG docker $USER
   ```

   После этого выйдите из системы и зайдите снова, чтобы изменения вступили в силу.

### Шаг 3: Установка Docker Compose

1. **Загрузите последнюю версию Docker Compose**:

   ```bash
   sudo curl -L "https://github.com/docker/compose/releases/latest/download/docker-compose-$(uname -s)-$(uname -m)" -o /usr/local/bin/docker-compose
   ```

2. **Сделайте Docker Compose исполняемым файлом**:

   ```bash
   sudo chmod +x /usr/local/bin/docker-compose
   ```

3. **Проверьте установку**:

   ```bash
   docker-compose --version
   ```

### Шаг 4: Проверка установки Docker и Docker Compose

Для проверки работы Docker создайте тестовый контейнер:

```bash
docker run hello-world
```

Если видите сообщение "Hello from Docker!", значит, Docker установлен правильно.

На этом установка Docker и Docker Compose завершена!
