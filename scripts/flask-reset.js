const isVenv = !!process.env.VIRTUAL_ENV;

if (isVenv) {
  console.log("Python virtual environment is active:", process.env.VIRTUAL_ENV, '\n\n Run \"deactivate\" to use flask-reset\n');
} else {
  const { exec } = require("child_process");

// Helper to run commands in sequence
function runCommand(cmd, options = {}) {
  return new Promise((resolve, reject) => {
    const process = exec(cmd, { shell: true, ...options }, (error, stdout, stderr) => {
      if (error) {
        reject(error);
        return;
      }
      resolve({ stdout, stderr });
    });

    process.stdout.on("data", (data) => process.stdout.write(data));
    process.stderr.on("data", (data) => process.stderr.write(data));
  });
}

async function resetVenv() {
  try {
    console.log("Removing old venv...");
    await runCommand('rmdir /S /Q .venv'); // remove old venv

    console.log("Creating new venv...");
    await runCommand('py -3 -m venv .venv'); // create new venv
    
    console.log("Installing Flask...");
    // Use the venv's pip directly
    await runCommand('.venv\\Scripts\\pip install Flask');
    
    console.log("Done! New venv created and Flask installed.");
    console.log("Activate it with: .venv\\Scripts\\activate");
    
    console.log("Removing __pycache__...");
    await runCommand('rmdir /S /Q __pycache__'); // remove Python cache

} catch (err) {
    console.error("Error:", err);
  }
}

resetVenv();
}