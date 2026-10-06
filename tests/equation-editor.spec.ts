import { expect, test } from '@playwright/test';

test('edits an existing equation and reloads the patched DOCX', async ({ page }) => {
  test.setTimeout(180_000);
  const errors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  page.on('pageerror', (error) => errors.push(String(error)));

  await page.goto('/');
  await expect(page.getByText('30 equations ready to edit')).toBeVisible({ timeout: 120_000 });
  await expect(page.getByRole('button', { name: 'Export DOCX' })).toBeEnabled({ timeout: 120_000 });

  const equations = page.locator('.editor-host math');
  await expect(equations).toHaveCount(17, { timeout: 120_000 });
  await page.evaluate(() => window.scrollTo(0, 650));
  await equations.nth(7).click();

  const dialog = page.getByRole('dialog', { name: 'Edit LaTeX' });
  await expect(dialog).toBeVisible();
  const input = page.getByLabel('LaTeX expression');
  await input.fill(String.raw`\frac{1}{2}`);
  await expect(page.getByTestId('equation-preview')).toContainText('12');
  const scrollBeforeUpdate = await page.evaluate(() => window.scrollY);
  await page.getByRole('button', { name: 'Done' }).click();

  await expect(page.getByText('30 equations ready to edit')).toBeVisible({ timeout: 120_000 });
  await expect(page.getByRole('button', { name: 'Export DOCX' })).toBeEnabled({ timeout: 120_000 });
  await expect.poll(async () => Math.abs((await page.evaluate(() => window.scrollY)) - scrollBeforeUpdate)).toBeLessThan(8);
  await expect(equations.nth(7)).toBeVisible({ timeout: 120_000 });
  await equations.nth(7).click();
  await expect(input).toHaveValue(String.raw`\frac{1}{2}`);
  expect(errors).toEqual([]);
});
