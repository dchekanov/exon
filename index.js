const childProcess = require('child_process');
const fsp = require('fs').promises;

class Exon {
  /**
   * @param {Object} [params]
   */
  constructor(params = {}) {
    ['appId', 'certificateFilePath', 'organization'].forEach(k => {
      if (!params.connection?.[k]) {
        throw new Error(`connection.${k} parameter is missing`);
      }
    });

    this.params = params;
    this.isBusy = false;
    this.isConnected = false;
    this.log('init', params);
  }

  /**
   * Expression to detect PS command prompt in stdout.
   * @returns {RegExp}
   */
  get promptRe() {
    return /^PS .*> $/;
  }

  /**
   * Log debug data.
   * @param {string} event
   * @param {*} data
   */
  log(event, data) {
    if (this.params.logFilePath) {
      data = typeof data === 'object' ? JSON.stringify(data) : data;

      fsp.appendFile(this.params.logFilePath, `${new Date().toISOString()} ${event} ${data ?? ''}\n`)
        .catch(console.error);
    }
  }

  /**
   * Generate cmdlet string.
   * @param {string} cmdletName
   * @param {Object} cmdletParams
   * @returns {string}
   */
  createCmdlet(cmdletName, cmdletParams) {
    let cmdlet = cmdletName;

    Object.entries(cmdletParams || {}).forEach(([paramName, paramValue]) => {
      cmdlet += ' -' + paramName.substr(0, 1).toUpperCase() + paramName.substr(1);

      if (typeof paramValue === 'boolean') {
        cmdlet += ':$' + paramValue.toString();
      } else {
        cmdlet += ` '${paramValue.replace(/'/g, '\'\'')}'`;
      }
    });

    return cmdlet;
  }

  /**
   * Spawn PS process.
   * @returns {Promise}
   */
  spawnPs() {
    this.log('spawn-ps-start');

    return new Promise((resolve, reject) => {
      let stdout = '';

      // reject if it takes too long
      const timeout = setTimeout(() => {
        this.log('spawn-ps-timeout', stdout);
        this.detachPsEventHandlers();
        reject(new Error('Timeout'));
      }, this.params.timeouts?.spawn || 5000);

      this.ps = childProcess.spawn('pwsh');

      this.ps.stdout.on('data', data => {
        stdout += data.toString();

        if (this.promptRe.test(data.toString())) {
          this.log('spawn-ps-success', stdout);
          this.detachPsEventHandlers();
          clearTimeout(timeout);
          resolve();
        }
      });

      this.ps.on('error', err => {
        this.log('spawn-ps-failure', err.message);
        this.detachPsEventHandlers();
        clearTimeout(timeout);
        reject(err);
      });

      this.ps.on('exit', code => {
        this.log('ps-exit', {code});
        this.ps = null;
      });
    });
  }

  /**
   * Stop listening for stdout/stderr and process errors once the outcome of an operation is known.
   */
  detachPsEventHandlers() {
    this.ps.stdout.removeAllListeners('data');
    this.ps.stderr.removeAllListeners('data');
    this.ps.removeAllListeners('error');
  }

  /**
   * Kill PS process.
   * @returns {Promise}
   */
  killPs() {
    this.log('kill-ps-start');

    return new Promise((resolve, reject) => {
      if (!this.ps) {
        this.log('kill-ps-success');
        return resolve();
      }

      this.ps.on('error', err => {
        this.log('kill-ps-failure', err.message);
        this.detachPsEventHandlers();
        reject(err);
      });

      if (this.ps.kill()) {
        this.log('kill-ps-success');
        // otherwise it sets this.ps to null after this promise is resolved
        this.ps.removeAllListeners('exit');
        this.ps = null;
        this.isConnected = false;
        this.isBusy = false;
        resolve();
      }
    });
  }

  /**
   * Execute a cmdlet.
   * @param {string} cmdletName
   * @param {Object} [cmdletParams]
   * @param {Object} [execParams]
   * @returns {Promise}
   */
  async exec(cmdletName, cmdletParams = {}, execParams = {}) {
    if (!this.ps) {
      await this.spawnPs();

      return this.exec(cmdletName, cmdletParams);
    }

    if (!this.isConnected && cmdletName !== 'Connect-ExchangeOnline') {
      await this.connect();

      return this.exec(cmdletName, cmdletParams);
    }

    this.log('exec-start', {cmdletName, cmdletParams});

    return new Promise((resolve, reject) => {
      if (this.isBusy) {
        this.log('exec-busy');
        return reject(new Error('Busy. Wait for the previous command to finish.'));
      }

      this.isBusy = true;

      let stdout = '';
      let stderr = '';

      // reject if it takes too long
      const timeout = setTimeout(() => {
        this.log('exec-timeout', {stdout, stderr});
        cleanup();
        reject(new Error('Timeout'));
      }, this.params.timeouts?.exec || 180000);

      /**
       * Cleanup after command has been executed.
       */
      const cleanup = () => {
        this.detachPsEventHandlers();
        this.isBusy = false;
        clearTimeout(timeout);
      };

      const cmdlet = this.createCmdlet(cmdletName, cmdletParams);

      this.ps.stdout.on('data', data => {
        if (this.promptRe.test(data.toString())) {
          cleanup();

          if (stderr) {
            this.log('exec-failure', stderr);

            // "ERROR_WSMAN_INVALID_SELECTORS" seem to be resolvable by trying again after some time
            if (stderr.toString().indexOf('ERROR_WSMAN_INVALID_SELECTORS') !== -1) {
              resolve(
                new Promise(resolve => setTimeout(resolve, 60000)).then(() => this.exec(cmdletName, cmdletParams))
              );
            } else {
              reject(new Error(stderr));
            }
          } else {
            this.log('exec-success', execParams.logStdout === false ? '' : stdout);
            resolve(stdout);
          }
        } else {
          stdout += data.toString();
        }
      });

      this.ps.stderr.on('data', data => {
        stderr += data.toString();
      });

      this.ps.stdin.write(`${cmdlet}\n`);
    });
  }

  /**
   * Start Exchange session.
   * @returns {Promise}
   */
  async connect() {
    await this.exec(
      'Connect-ExchangeOnline',
      {
        ...this.params.connection,
        ShowBanner: false
      },
      {
        logStdout: false
      }
    );

    this.isConnected = true;
  }

  /**
   * End Exchange session.
   * @returns {Promise}
   */
  async disconnect() {
    if (!this.ps || !this.isConnected) {
      return;
    }

    await this.exec('Disconnect-ExchangeOnline', {confirm: false});

    this.isBusy = false;
    this.isConnected = false;
  }

  /**
   * Disconnect from Exchange and kill PS process.
   * @returns {Promise<void>}
   */
  async reset() {
    if (!this.isBusy) {
      await this.disconnect();
    }

    await this.killPs();
  }
}

module.exports = Exon;
