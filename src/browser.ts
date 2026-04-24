import type { Browser } from 'puppeteer-core';
import puppeteer from 'puppeteer-core';

import { config } from './config.js';
import { logger } from './logger.js';

let browserPromise: Promise<Browser> | null = null;

const launchArgs = [
  '--no-sandbox',
  '--disable-setuid-sandbox',
  '--disable-dev-shm-usage',
  '--disable-accelerated-2d-canvas',
  '--disable-gpu',
  '--no-first-run',
  '--no-zygote',
  '--hide-scrollbars',
  '--disable-background-timer-throttling',
  '--disable-backgrounding-occluded-windows',
  '--disable-renderer-backgrounding',
];

const resolveExecutablePath = (): string => {
  if (config.chromiumExecutablePath) {
    return config.chromiumExecutablePath;
  }
  throw new Error(
    'CHROMIUM_EXECUTABLE_PATH is not set. In Docker, this is baked in (/usr/bin/chromium). In dev local, configure it in .env.local.',
  );
};

const launch = async (): Promise<Browser> => {
  const executablePath = resolveExecutablePath();
  logger.info({ executablePath }, 'Launching Chromium');
  const browser = await puppeteer.launch({
    executablePath,
    headless: true,
    args: launchArgs,
    defaultViewport: { width: 1280, height: 800, deviceScaleFactor: 1 },
  });
  browser.on('disconnected', () => {
    logger.warn('Chromium disconnected — will relaunch on next request');
    browserPromise = null;
  });
  return browser;
};

export const getBrowser = async (): Promise<Browser> => {
  if (!browserPromise) {
    browserPromise = launch().catch((err) => {
      browserPromise = null;
      throw err;
    });
  }
  return browserPromise;
};

export const warmUpBrowser = async (): Promise<void> => {
  try {
    await getBrowser();
    logger.info('Browser warmed up');
  } catch (err) {
    logger.error({ err }, 'Browser warm-up failed');
  }
};

export const closeBrowser = async (): Promise<void> => {
  if (!browserPromise) return;
  try {
    const browser = await browserPromise;
    await browser.close();
  } catch (err) {
    logger.warn({ err }, 'Error closing browser');
  } finally {
    browserPromise = null;
  }
};
