async function executeCode({ language, code }) {
  // Валидация входных данных
  if (!['javascript', 'python'].includes(language)) {
    throw new Error('Unsupported language');
  }

  if (language === 'javascript') {
    try {
      // Для JavaScript используем vm2 для безопасного выполнения
      const { VM } = require('vm2');
      const vm = new VM({
        timeout: 5000,
        sandbox: {}
      });
      
      return {
        success: true,
        output: vm.run(code)
      };
    } catch (error) {
      return {
        success: false,
        error: error.message
      };
    }
  }

  if (language === 'python') {
    try {
      const { spawnSync } = require('child_process');
      const result = spawnSync('python', ['-c', code], {
        timeout: 5000,
        encoding: 'utf-8'
      });

      if (result.error) {
        throw result.error;
      }

      return {
        success: true,
        output: result.stdout,
        error: result.stderr
      };
    } catch (error) {
      return {
        success: false,
        error: error.message
      };
    }
  }
}

module.exports = executeCode; 