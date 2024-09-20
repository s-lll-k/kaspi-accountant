module.exports = {
    apps: [
        {
            name: 'my-telegram-bot', // Имя приложения
            script: './index.js',    // Путь к вашему основному скрипту
            instances: 1,            // Количество экземпляров
            autorestart: true,       // Автоматический перезапуск при сбое
            watch: false,            // Слежение за изменениями файлов (можно включить, если нужно)
            max_memory_restart: '1G', // Перезапуск при превышении объема памяти
            env: {
            NODE_ENV: 'development',
            PORT: 3000
            },
            env_production: {
             NODE_ENV: 'production'
            }
        }
    ]
};