import { expect, test } from '@playwright/test';

test('edits an existing equation in place', async ({ page }) => {
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
  await page.locator('.superdoc').evaluate((element) => element.setAttribute('data-edit-marker', 'preserved'));
  await page.getByRole('button', { name: 'Done' }).click();

  await expect(page.getByText('Equation 8 updated in place with SuperDoc')).toBeVisible({ timeout: 120_000 });
  await expect(page.locator('.superdoc[data-edit-marker="preserved"]')).toHaveCount(1);
  await expect(page.getByRole('button', { name: 'Export DOCX' })).toBeEnabled({ timeout: 120_000 });
  await expect.poll(async () => Math.abs((await page.evaluate(() => window.scrollY)) - scrollBeforeUpdate)).toBeLessThan(8);
  await expect(equations.nth(7)).toBeVisible({ timeout: 120_000 });
  await equations.nth(7).click();
  await expect(input).toHaveValue(String.raw`\frac{1}{2}`);
  await page.getByRole('button', { name: 'Cancel' }).click();
  await page.locator('.superdoc-text-run[styleid="Heading1"]').filter({ hasText: 'Limits and derivatives' }).click();
  await page.waitForTimeout(500);
  const editedWidth = await equations.nth(7).evaluate((equation) => equation.getBoundingClientRect().width);
  await equations.nth(7).click({ position: { x: editedWidth - 1, y: 2 } });
  await expect(input).toHaveValue(String.raw`\frac{1}{2}`);
  await page.waitForTimeout(500);
  expect(errors).toEqual([]);
});

test('inserts a new equation at the current cursor', async ({ page }) => {
  test.setTimeout(180_000);
  const errors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  page.on('pageerror', (error) => errors.push(String(error)));

  await page.goto('/');
  await expect(page.getByText('30 equations ready to edit')).toBeVisible({ timeout: 120_000 });

  const title = page.locator('.superdoc-text-run[styleid="Title"]');
  await expect(title).toHaveCount(1, { timeout: 120_000 });
  await title.click({ position: { x: 180, y: 20 } });
  await page.getByRole('button', { name: 'Insert equation' }).click();

  const dialog = page.getByRole('dialog', { name: 'Insert equation' });
  await expect(dialog).toBeVisible();
  const input = page.getByLabel('LaTeX expression');
  await expect(input).toHaveValue('');
  await page.locator('.superdoc').evaluate((element) => element.setAttribute('data-noop-marker', 'preserved'));
  await page.getByRole('button', { name: 'Done' }).click();
  await expect(dialog).toBeHidden();
  await expect(page.locator('.superdoc[data-noop-marker="preserved"]')).toHaveCount(1);
  await expect(page.getByText('30 equations ready to edit')).toBeVisible();

  await title.click({ position: { x: 180, y: 20 } });
  await page.getByRole('button', { name: 'Insert equation' }).click();
  await input.fill(String.raw`\sum_{n=1}^{10} n`);
  await expect(page.getByTestId('equation-preview').locator('.katex')).toBeVisible();
  await page.locator('.superdoc').evaluate((element) => element.setAttribute('data-insertion-marker', 'preserved'));
  await page.getByRole('button', { name: 'Done' }).click();

  await expect(page.getByText('Equation inserted in place with SuperDoc')).toBeVisible({ timeout: 120_000 });
  await expect(page.locator('.superdoc[data-insertion-marker="preserved"]')).toHaveCount(1);
  await expect(page.locator('.superdoc-fragment[styleid="Title"] math')).toHaveCount(1);
  expect(errors).toEqual([]);
});

test('deletes an existing equation when blank LaTeX is submitted', async ({ page }) => {
  test.setTimeout(180_000);
  const errors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  page.on('pageerror', (error) => errors.push(String(error)));

  await page.goto('/');
  await expect(page.getByText('30 equations ready to edit')).toBeVisible({ timeout: 120_000 });
  const equations = page.locator('.editor-host math');
  await expect(equations).toHaveCount(17, { timeout: 120_000 });
  const sourceNodeId = await equations.nth(7).evaluate(
    (equation) => equation.closest<HTMLElement>('.superdoc-fragment')?.dataset.sourceNodeId,
  );
  expect(sourceNodeId).toBeTruthy();
  await equations.nth(7).click();

  const dialog = page.getByRole('dialog', { name: 'Edit LaTeX' });
  await expect(dialog).toBeVisible();
  await page.getByLabel('LaTeX expression').fill('');
  await page.getByRole('button', { name: 'Delete' }).click();

  await expect(page.getByText('Equation 8 deleted in place with SuperDoc')).toBeVisible({ timeout: 120_000 });
  await expect(dialog).toBeHidden();
  const updatedFragment = page.locator(`.superdoc-fragment[data-source-node-id="${sourceNodeId}"]`);
  await expect(updatedFragment).toHaveCount(1, { timeout: 120_000 });
  await expect(updatedFragment.locator('math')).toHaveCount(0);
  expect(errors).toEqual([]);
});

test('places the caret after a clicked equation for atomic Backspace deletion', async ({ page }) => {
  test.setTimeout(180_000);
  const errors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  page.on('pageerror', (error) => errors.push(String(error)));

  await page.goto('/');
  await expect(page.getByText('30 equations ready to edit')).toBeVisible({ timeout: 120_000 });
  const equations = page.locator('.editor-host math');
  await expect(equations).toHaveCount(17, { timeout: 120_000 });
  const sourceNodeId = await equations.nth(7).evaluate(
    (equation) => equation.closest<HTMLElement>('.superdoc-fragment')?.dataset.sourceNodeId,
  );
  expect(sourceNodeId).toBeTruthy();
  await equations.nth(7).click();
  await expect(page.getByRole('dialog', { name: 'Edit LaTeX' })).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.locator('textarea[data-v2-ime-host="true"]')).toBeFocused();
  const caret = page.locator('[data-v2-local-selection-caret="true"]');
  await expect(caret).toBeVisible();
  const equationBox = await equations.nth(7).boundingBox();
  expect(equationBox).not.toBeNull();
  await expect.poll(async () => (await caret.boundingBox())?.x ?? -1)
    .toBeGreaterThanOrEqual(equationBox!.x + equationBox!.width - 2);
  await page.keyboard.press('ArrowLeft');
  await expect.poll(async () => (await caret.boundingBox())?.x ?? Number.POSITIVE_INFINITY)
    .toBeLessThanOrEqual(equationBox!.x + 2);
  await page.keyboard.press('ArrowRight');
  await expect.poll(async () => (await caret.boundingBox())?.x ?? -1)
    .toBeGreaterThanOrEqual(equationBox!.x + equationBox!.width - 2);
  await page.keyboard.press('Backspace');

  await expect(page.locator(`.superdoc-fragment[data-source-node-id="${sourceNodeId}"] math`))
    .toHaveCount(0, { timeout: 120_000 });
  expect(errors).toEqual([]);
});
