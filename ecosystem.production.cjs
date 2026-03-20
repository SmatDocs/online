const blueRoot = process.env.PRODUCTION_BLUE_ROOT || "/home/humphry/online";
const greenRoot = process.env.PRODUCTION_GREEN_ROOT || "/home/humphry/online-green";

function makeApp(name, root, port) {
    return {
        name,
        cwd: root,
        script: "./start-coolwsd.sh",
        interpreter: "bash",
        exec_mode: "fork",
        watch: false,
        autorestart: true,
        restart_delay: 1000,
        kill_timeout: 15000,
        time: true,
        env: {
            COOLWSD_CONFIG_FILE: `${root}/coolwsd_prod.xml`,
            COOLWSD_PORT: String(port),
            COOLWSD_SYS_TEMPLATE_PATH: `${root}/systemplate`,
            COOLWSD_CHILD_ROOT_PATH: `${root}/jails`,
            COOLWSD_CACHE_PATH: `${root}/cache`,
        },
    };
}

module.exports = {
    apps: [
        makeApp(
            process.env.PRODUCTION_BLUE_PM2_NAME || "coolwsd-blue",
            blueRoot,
            process.env.PRODUCTION_BLUE_PORT || 9980
        ),
        makeApp(
            process.env.PRODUCTION_GREEN_PM2_NAME || "coolwsd-green",
            greenRoot,
            process.env.PRODUCTION_GREEN_PORT || 9981
        ),
    ],
};
