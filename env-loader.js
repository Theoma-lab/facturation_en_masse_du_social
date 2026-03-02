/**
 * Simple .env loader for browser environments without a build step.
 * Note: This is less secure than a build-time injection but keeps secrets out of the main code.
 */
async function loadEnv() {
    console.log('env-loader: Fetching env.local.txt...');
    try {
        const response = await fetch('env.local.txt');
        if (!response.ok) {
            console.warn('env-loader: env.local.txt file not found or browser blocked fetch.');
            throw new Error('Could not load configuration file');
        }

        const text = await response.text();
        console.log('env-loader: .env content retrieved.');
        const env = {};

        text.split(/\r?\n/).forEach(line => {
            const trimmedLine = line.trim();
            if (!trimmedLine || trimmedLine.startsWith('#')) return;

            const index = trimmedLine.indexOf('=');
            if (index > -1) {
                const key = trimmedLine.substring(0, index).trim();
                const value = trimmedLine.substring(index + 1).trim();
                if (key) env[key] = value;
            }
        });

        return env;
    } catch (error) {
        console.error('Error loading environment variables:', error);
        return {};
    }
}

window.loadEnv = loadEnv;
