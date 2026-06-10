import { test, expect } from '@playwright/test';

test('test', async ({ page }) => {
  await page.goto('https://web.whatsapp.com/');
  await page.getByRole('button', { name: 'Close' }).click();
  await page.getByTestId('message-yourself-row').getByText('Akash').click();
  await page.getByRole('button', { name: 'Attach' }).click();
  await page.getByRole('menuitem', { name: 'Photos & videos' }).click();
  await page.getByTestId('conversation-compose-box-input').setInputFiles('testing.png');
  await page.getByRole('button', { name: 'Send 1 selected' }).click();
});