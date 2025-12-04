const { spawn } = require("child_process");

const child = spawn("npm", ["start"], {
  shell: true,
  stdio: "inherit"
});

child.on("close", (code) => {
  console.log(`Process exited with code ${code}`);
});
