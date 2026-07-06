import type { Browser } from 'puppeteer-core';
import puppeteer from 'puppeteer-core';
import chromium from '@sparticuz/chromium';

import { config } from './config.js';
import { logger } from './logger.js';

let browserPromise: Promise<Browser> | null = null;

// Usado só em dev local, com CHROMIUM_EXECUTABLE_PATH apontando pro Chrome do sistema.
const LOCAL_LAUNCH_ARGS = [
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

type LaunchConfig = { executablePath: string; headless: true | 'shell'; args: string[] };

const resolveLaunchConfig = async (): Promise<LaunchConfig> => {
  if (config.chromiumExecutablePath) {
    return { executablePath: config.chromiumExecutablePath, headless: true, args: LOCAL_LAUNCH_ARGS };
  }
  // @sparticuz/chromium roda em sandboxes restritos (ex.: gVisor) sem crashar — o Chromium
  // do apt trava com SIGTRAP ao fazer qualquer request de rede nesse tipo de ambiente.
  return {
    executablePath: await chromium.executablePath(),
    headless: 'shell',
    // Filtra --font-render-hinting=none: o default deles muda a renderização de texto do PDF.
    args: puppeteer
      .defaultArgs({ args: chromium.args, headless: 'shell' })
      .filter((arg) => arg !== '--font-render-hinting=none'),
  };
};

const launch = async (): Promise<Browser> => {
  const { executablePath, headless, args } = await resolveLaunchConfig();
  logger.info({ executablePath }, 'Launching Chromium');
  const browser = await puppeteer.launch({
    executablePath,
    headless,
    args,
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
