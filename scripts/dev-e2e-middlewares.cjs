/**
 * Dev-only E2E / coding-agent middleware endpoints (EXTRACTED from webpack.config.cjs).
 *
 * These endpoints implement the E2E-test logging, loop-control, test-case CRUD,
 * and prompt CRUD APIs used by the external E2E/coding-agent harness. They are
 * registered ONLY when the environment variable ENABLE_DEV_ENDPOINTS=true is set
 * (see webpack.config.cjs); they are never part of the production add-in build
 * or the production server (scripts/docker-server.cjs).
 *
 * Security notes:
 * - All endpoints intentionally use `Access-Control-Allow-Origin: *` because the
 *   harness posts from the add-in origin during local development. This is why
 *   the module is flag-gated and why the dev server binds 127.0.0.1 by default.
 * - Several endpoints WRITE files under the project root (logs/, e2e-test-cases-
 *   dynamic.json, user-prompts.json). Do not enable on an exposed interface.
 *
 * Signal handlers use process.once() so repeated setup cannot stack duplicates
 * and foreign SIGINT/SIGTERM handlers are never wiped.
 *
 * @param {Array} middlewares - webpack-dev-server middlewares (passed through)
 * @param {object} devServer - webpack-dev-server instance
 * @returns {Array} the middlewares, unchanged
 */

const path = require('path');
const fs = require('fs');

// Paths are resolved against the project root (this file lives in scripts/),
// matching the layout the harness expects (logs/, prompts.json, ... at root).
const rootDir = path.resolve(__dirname, '..');

function setupDevE2eMiddlewares(middlewares, devServer) {
  if (!devServer) {
    throw new Error('webpack-dev-server is not defined');
  }

  const express = require('express');
  const app = devServer.app;

  // Enable JSON body parsing for API routes
  app.use('/api/prompts', express.json());
  app.use('/log', express.json());
  app.use('/api/test-cases', express.json());
  app.use('/api/trace-log', express.json());
  app.use('/api/fix-log', express.json());

  // ============================================================================
  // E2E TEST LOGGING ENDPOINTS
  // ============================================================================

  // Log file path for persistent storage
  const logsDir = path.join(rootDir, 'logs');
  const logFilePath = path.join(logsDir, 'e2e-test-logs.json');

  // Ensure logs directory exists
  if (!fs.existsSync(logsDir)) {
    fs.mkdirSync(logsDir, { recursive: true });
  }

  // Load existing logs from file on server start
  let logs = [];
  try {
    if (fs.existsSync(logFilePath)) {
      const fileContent = fs.readFileSync(logFilePath, 'utf8');
      if (fileContent.trim()) {
        logs = JSON.parse(fileContent);
        console.log(`[E2E Logs] Loaded ${logs.length} existing log entries from ${logFilePath}`);
      }
    }
  } catch (error) {
    console.error(`[E2E Logs] Error loading logs from file: ${error.message}`);
    logs = []; // Start fresh if file is corrupted
  }

  global.e2eLogs = logs; // Make accessible to external scripts

  // Function to persist logs to file
  const persistLogs = () => {
    try {
      // Write entire logs array to file (append mode would be complex with JSON)
      fs.writeFileSync(logFilePath, JSON.stringify(logs, null, 2), 'utf8');
    } catch (error) {
      console.error(`[E2E Logs] Error persisting logs to file: ${error.message}`);
    }
  };

  // Persist logs periodically (every 10 entries) and on server shutdown
  let logWriteCounter = 0;
  const LOG_WRITE_INTERVAL = 10;

  // POST /log - Receive logs from Word add-in
  app.post('/log', (req, res) => {
    const logEntry = req.body;
    logEntry.receivedAt = new Date().toISOString();
    logs.push(logEntry);

    // Print to console for immediate visibility
    if (logEntry.message) {
      console.log(`[Client Log] ${logEntry.message}`);
    } else {
      console.log('[Client Log]', logEntry);
    }

    // Persist to file periodically (every LOG_WRITE_INTERVAL entries)
    logWriteCounter++;
    if (logWriteCounter >= LOG_WRITE_INTERVAL) {
      persistLogs();
      logWriteCounter = 0;
    }

    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.status(200).end();
  });

  // Note: SIGINT and SIGTERM handlers are set up after fix log initialization
  // to ensure both regular logs and fix logs are persisted on shutdown

  // OPTIONS /log - CORS preflight
  app.options('/log', (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.status(204).end();
  });

  // GET /logs - Retrieve logs (for coding assistant)
  app.get('/logs', (req, res) => {
    const since = req.query.since ? new Date(req.query.since) : null;
    const filtered = since
      ? logs.filter(log => new Date(log.timestamp) > since)
      : logs.slice(-1000); // Last 1000 by default

    res.setHeader('Content-Type', 'application/json');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.json(filtered);
  });

  // POST /logs/clear - Clear logs (both memory and file)
  app.post('/logs/clear', (req, res) => {
    logs.length = 0;
    try {
      // Clear the log file as well
      fs.writeFileSync(logFilePath, '[]', 'utf8');
      console.log('[E2E Logs] Cleared logs from memory and file');
    } catch (error) {
      console.error(`[E2E Logs] Error clearing log file: ${error.message}`);
    }
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.json({ success: true, message: 'Logs cleared from memory and file' });
  });

  // ============================================================================
  // TRACE LOG API (Record & Replay)
  // ============================================================================

  // POST /api/trace-log - Receive and store trace logs from Stability Loop failures
  app.post('/api/trace-log', (req, res) => {
    const traceData = req.body;
    console.log(`[Trace Log] Received trace for test run ${traceData.testRunNumber}`);

    try {
      // Ensure logs directory exists
      const logsDir = path.join(rootDir, 'logs');
      if (!fs.existsSync(logsDir)) {
        fs.mkdirSync(logsDir, { recursive: true });
      }

      // Save trace to file: logs/trace-log-{testRunNumber}.json
      const traceFileName = `trace-log-${traceData.testRunNumber}.json`;
      const traceFilePath = path.join(logsDir, traceFileName);

      // Trace format includes:
      // - testRunNumber: Test run number
      // - testId: Unique test ID
      // - originalText: Original document text (for test setup)
      // - expectedText: Expected LLM output (for verification)
      // - finalText: Final document text (for comparison)
      // - trace: Array of trace entries with exact API calls
      // - timestamp: When trace was saved
      fs.writeFileSync(traceFilePath, JSON.stringify(traceData, null, 2), 'utf8');

      console.log(`[Trace Log] Saved trace to ${traceFilePath} (${traceData.trace?.length || 0} entries)`);

      res.setHeader('Access-Control-Allow-Origin', '*');
      res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
      res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
      res.json({
        success: true,
        message: `Trace saved to ${traceFileName}`,
        traceLength: traceData.trace?.length || 0
      });
    } catch (error) {
      console.error(`[Trace Log] Error saving trace: ${error.message}`);
      res.setHeader('Access-Control-Allow-Origin', '*');
      res.status(500).json({
        success: false,
        error: `Failed to save trace: ${error.message}`
      });
    }
  });

  // OPTIONS /api/trace-log - CORS preflight
  app.options('/api/trace-log', (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.status(204).end();
  });

  // ============================================================================
  // FIX LOG API (Separate from E2E logs)
  // ============================================================================

  // Fix log file path for persistent storage (reuse logsDir from above)
  const fixLogFilePath = path.join(logsDir, 'fix-logs.json');

  // Load existing fix logs from file on server start
  let fixLogs = [];
  try {
    if (fs.existsSync(fixLogFilePath)) {
      const fileContent = fs.readFileSync(fixLogFilePath, 'utf8');
      if (fileContent.trim()) {
        fixLogs = JSON.parse(fileContent);
        console.log(`[Fix Logs] Loaded ${fixLogs.length} existing fix log entries from ${fixLogFilePath}`);
      }
    } else {
      // Initialize empty array if file doesn't exist
      fs.writeFileSync(fixLogFilePath, '[]', 'utf8');
      console.log(`[Fix Logs] Created new fix log file: ${fixLogFilePath}`);
    }
  } catch (error) {
    console.error(`[Fix Logs] Error loading fix logs from file: ${error.message}`);
    fixLogs = []; // Start fresh if file is corrupted
  }

  // Function to persist fix logs to file
  const persistFixLogs = () => {
    try {
      fs.writeFileSync(fixLogFilePath, JSON.stringify(fixLogs, null, 2), 'utf8');
    } catch (error) {
      console.error(`[Fix Logs] Error persisting fix logs to file: ${error.message}`);
    }
  };

  // POST /api/fix-log - Receive and store fix logs
  app.post('/api/fix-log', (req, res) => {
    const fixEntry = req.body;
    fixEntry.receivedAt = new Date().toISOString();
    fixLogs.push(fixEntry);

    // Persist immediately (fix logs are important and should be saved right away)
    persistFixLogs();

    console.log(`[Fix Log] Logged fix: ${fixEntry.metadata?.file || 'unknown'} - ${fixEntry.metadata?.issue?.substring(0, 50) || 'unknown'}...`);

    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.json({
      success: true,
      message: 'Fix logged successfully',
      totalFixes: fixLogs.length
    });
  });

  // OPTIONS /api/fix-log - CORS preflight
  app.options('/api/fix-log', (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.status(204).end();
  });

  // Persist both regular logs and fix logs on server shutdown.
  // process.once() keeps repeated setup from stacking duplicates without
  // wiping SIGINT/SIGTERM handlers installed by anyone else.
  const persistOnShutdown = () => {
    console.log('[E2E Logs] Persisting logs before shutdown...');
    persistLogs();
    console.log('[Fix Logs] Persisting fix logs before shutdown...');
    persistFixLogs();
    process.exit(0);
  };

  process.once('SIGINT', persistOnShutdown);
  process.once('SIGTERM', persistOnShutdown);

  // ============================================================================
  // E2E TEST LOOP CONTROL API
  // ============================================================================

  // Global state for loop control
  // CRITICAL: Initial state is PAUSED - loop will only proceed when explicitly triggered
  // This ensures the coding agent has time to analyze logs and fix issues before each iteration
  global.e2eLoopControl = global.e2eLoopControl || {
    canProceed: false,  // Start paused - coding agent must trigger each iteration
    waitingForTrigger: true,
    lastIteration: null
  };

  // OPTIONS /api/e2e-loop/status - CORS preflight
  app.options('/api/e2e-loop/status', (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.status(200).end();
  });

  // GET /api/e2e-loop/status - Check if loop can proceed
  app.get('/api/e2e-loop/status', (req, res) => {
    const control = global.e2eLoopControl;
    res.setHeader('Content-Type', 'application/json');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    // If canProceed is false, we're definitely waiting for trigger
    // This handles the case where pause() was called but server was unreachable
    const isWaiting = !control.canProceed || control.waitingForTrigger;

    res.json({
      canProceed: control.canProceed,
      waitingForTrigger: isWaiting,  // Ensure this reflects actual waiting state
      lastIteration: control.lastIteration
    });
  });

  // POST /api/e2e-loop/trigger - Trigger next iteration
  app.post('/api/e2e-loop/trigger', (req, res) => {
    const control = global.e2eLoopControl;
    control.canProceed = true;
    control.waitingForTrigger = false;
    res.setHeader('Content-Type', 'application/json');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.json({
      success: true,
      message: 'Loop trigger activated',
      canProceed: control.canProceed
    });
    console.log('[API] E2E loop trigger activated');
  });

  // POST /api/e2e-loop/pause - Pause the loop (set waiting state)
  app.post('/api/e2e-loop/pause', (req, res) => {
    const control = global.e2eLoopControl;
    control.canProceed = false;
    control.waitingForTrigger = true;
    res.setHeader('Content-Type', 'application/json');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.json({
      success: true,
      message: 'Loop paused, waiting for trigger',
      waitingForTrigger: true
    });
    console.log('[API] E2E loop paused');
  });

  // ============================================================================
  // E2E TEST CASES API
  // ============================================================================

  // GET /api/test-cases - Fetch test cases (static + dynamic)
  app.get('/api/test-cases', (req, res) => {
    console.log('[API] GET /api/test-cases');
    const staticTestsPath = path.join(rootDir, 'e2e-test-cases.json');
    const dynamicTestsPath = path.join(rootDir, 'e2e-test-cases-dynamic.json');

    try {
      // Load static test cases
      let staticTests = [];
      if (fs.existsSync(staticTestsPath)) {
        const staticTestsData = fs.readFileSync(staticTestsPath, 'utf8');
        staticTests = JSON.parse(staticTestsData);
      }

      // Load dynamic test cases
      let dynamicTests = [];
      if (fs.existsSync(dynamicTestsPath)) {
        const dynamicTestsData = fs.readFileSync(dynamicTestsPath, 'utf8');
        dynamicTests = JSON.parse(dynamicTestsData);
      }

      const allTests = [...staticTests, ...dynamicTests];

      res.setHeader('Content-Type', 'application/json');
      res.setHeader('Access-Control-Allow-Origin', '*');
      res.json(allTests);
      console.log(`[API] Sent ${allTests.length} test cases (${staticTests.length} static + ${dynamicTests.length} dynamic)`);
    } catch (error) {
      console.error('[API] Error reading test cases:', error.message);
      res.status(500).json({ error: 'Failed to load test cases' });
    }
  });

  // POST /api/test-cases - Add dynamic test cases
  app.post('/api/test-cases', (req, res) => {
    console.log('[API] POST /api/test-cases');
    const dynamicTestsPath = path.join(rootDir, 'e2e-test-cases-dynamic.json');

    try {
      const newTest = req.body;

      // Validate test case
      if (!newTest.original || !newTest.modified) {
        return res.status(400).json({ error: 'Invalid test case: missing required fields (original, modified)' });
      }

      // Load existing dynamic tests
      let dynamicTests = [];
      if (fs.existsSync(dynamicTestsPath)) {
        const dynamicTestsData = fs.readFileSync(dynamicTestsPath, 'utf8');
        dynamicTests = JSON.parse(dynamicTestsData);
      }

      // Add new test case
      const testCase = {
        id: newTest.id || `test-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        original: newTest.original,
        modified: newTest.modified,
        expected: newTest.modified, // Alias for modified
        reason: newTest.reason || 'auto-generated',
        createdAt: new Date().toISOString()
      };

      dynamicTests.push(testCase);

      // Save to file
      fs.writeFileSync(dynamicTestsPath, JSON.stringify(dynamicTests, null, 2), 'utf8');

      res.setHeader('Access-Control-Allow-Origin', '*');
      res.json({ success: true, testCase });
      console.log(`[API] Added test case: ${testCase.id}`);
    } catch (error) {
      console.error('[API] Error saving test case:', error.message);
      res.status(500).json({ error: 'Failed to save test case' });
    }
  });

  // ============================================================================

  // GET /api/prompts - Fetch all prompts (default + user custom)
  app.get('/api/prompts', (req, res) => {
    console.log('[API] GET /api/prompts');
    const defaultPromptsPath = path.join(rootDir, 'prompts.json');
    const userPromptsPath = path.join(rootDir, 'user-prompts.json');

    try {
      // Load default prompts
      const defaultPromptsData = fs.readFileSync(defaultPromptsPath, 'utf8');
      const defaultPrompts = JSON.parse(defaultPromptsData);

      // Load user custom prompts
      let userPrompts = [];
      if (fs.existsSync(userPromptsPath)) {
        const userPromptsData = fs.readFileSync(userPromptsPath, 'utf8');
        userPrompts = JSON.parse(userPromptsData);
      }

      // Merge: user prompts can override default prompts by ID
      const promptsMap = new Map();
      defaultPrompts.forEach(p => promptsMap.set(p.id, p));
      userPrompts.forEach(p => promptsMap.set(p.id, p));

      const allPrompts = Array.from(promptsMap.values());

      res.setHeader('Content-Type', 'application/json');
      res.setHeader('Access-Control-Allow-Origin', '*');
      res.json(allPrompts);
      console.log(`[API] Sent ${allPrompts.length} prompts (${defaultPrompts.length} default + ${userPrompts.length} user)`);
    } catch (error) {
      console.error('[API] Error reading prompts:', error.message);
      res.status(500).json({ error: 'Failed to load prompts' });
    }
  });

  // POST /api/prompts - Save/update a custom prompt
  app.post('/api/prompts', (req, res) => {
    console.log('[API] POST /api/prompts');
    const userPromptsPath = path.join(rootDir, 'user-prompts.json');

    try {
      const newPrompt = req.body;

      // Validate prompt
      if (!newPrompt.id || !newPrompt.name || !newPrompt.template) {
        return res.status(400).json({ error: 'Invalid prompt: missing required fields (id, name, template)' });
      }

      // Load existing user prompts
      let userPrompts = [];
      if (fs.existsSync(userPromptsPath)) {
        const userPromptsData = fs.readFileSync(userPromptsPath, 'utf8');
        userPrompts = JSON.parse(userPromptsData);
      }

      // Check if prompt exists (update) or is new (create)
      const existingIndex = userPrompts.findIndex(p => p.id === newPrompt.id);
      if (existingIndex >= 0) {
        userPrompts[existingIndex] = newPrompt;
        console.log(`[API] Updated prompt: ${newPrompt.id}`);
      } else {
        userPrompts.push(newPrompt);
        console.log(`[API] Created new prompt: ${newPrompt.id}`);
      }

      // Save to file
      fs.writeFileSync(userPromptsPath, JSON.stringify(userPrompts, null, 2), 'utf8');

      res.setHeader('Access-Control-Allow-Origin', '*');
      res.json({ success: true, prompt: newPrompt });
    } catch (error) {
      console.error('[API] Error saving prompt:', error.message);
      res.status(500).json({ error: 'Failed to save prompt' });
    }
  });

  // DELETE /api/prompts/:id - Delete a custom prompt
  app.delete('/api/prompts/:id', (req, res) => {
    const promptId = req.params.id;
    console.log(`[API] DELETE /api/prompts/${promptId}`);
    const userPromptsPath = path.join(rootDir, 'user-prompts.json');

    try {
      // Load existing user prompts
      let userPrompts = [];
      if (fs.existsSync(userPromptsPath)) {
        const userPromptsData = fs.readFileSync(userPromptsPath, 'utf8');
        userPrompts = JSON.parse(userPromptsData);
      }

      // Filter out the prompt to delete
      const filteredPrompts = userPrompts.filter(p => p.id !== promptId);

      if (filteredPrompts.length === userPrompts.length) {
        // Prompt not found in user prompts - might be a default prompt
        console.log(`[API] Prompt ${promptId} not found in user prompts (might be default)`);
        return res.status(404).json({ error: 'Prompt not found or cannot delete default prompt' });
      }

      // Save updated list
      fs.writeFileSync(userPromptsPath, JSON.stringify(filteredPrompts, null, 2), 'utf8');

      res.setHeader('Access-Control-Allow-Origin', '*');
      res.json({ success: true, message: `Deleted prompt: ${promptId}` });
      console.log(`[API] Deleted prompt: ${promptId}`);
    } catch (error) {
      console.error('[API] Error deleting prompt:', error.message);
      res.status(500).json({ error: 'Failed to delete prompt' });
    }
  });

  return middlewares;
}

module.exports = setupDevE2eMiddlewares;
