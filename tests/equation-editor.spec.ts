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
  await page.getByRole('button', { name: 'Done' }).click();

  await expect(page.getByText('31 equations ready to edit')).toBeVisible({ timeout: 120_000 });
  await expect(page.locator('.superdoc-fragment[styleid="Title"]')).toHaveCount(2, { timeout: 120_000 });
  const displayLayout = await page.evaluate(() => {
    const fragments = Array.from(document.querySelectorAll<HTMLElement>('.superdoc-fragment'));
    const before = fragments.find((fragment) => fragment.textContent?.trim() === 'Equation R');
    const after = fragments.find((fragment) => fragment.textContent?.trim() === 'endering Fixture');
    const pageElement = before?.closest<HTMLElement>('.superdoc-page');
    if (!before || !after || !pageElement) return null;
    const beforeBounds = before.getBoundingClientRect();
    const afterBounds = after.getBoundingClientRect();
    const equationFragment = fragments.find((fragment) => {
      const bounds = fragment.getBoundingClientRect();
      return Boolean(fragment.querySelector('math')) && bounds.top >= beforeBounds.bottom && bounds.bottom <= afterBounds.top;
    });
    const math = equationFragment?.querySelector('math');
    if (!equationFragment || !math) return null;
    const equationBounds = equationFragment.getBoundingClientRect();
    const mathBounds = math.getBoundingClientRect();
    const pageBounds = pageElement.getBoundingClientRect();
    return {
      centeredDifference: Math.abs(mathBounds.left + mathBounds.width / 2 - (pageBounds.left + pageBounds.width / 2)),
      hasSeparateRow: beforeBounds.bottom <= equationBounds.top && equationBounds.bottom <= afterBounds.top,
    };
  });
  expect(displayLayout?.hasSeparateRow).toBe(true);
  expect(displayLayout?.centeredDifference).toBeLessThan(2);
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

  await expect(page.getByText('29 equations ready to edit')).toBeVisible({ timeout: 120_000 });
  await expect(dialog).toBeHidden();
  const updatedFragment = page.locator(`.superdoc-fragment[data-source-node-id="${sourceNodeId}"]`);
  await expect(updatedFragment).toHaveCount(1, { timeout: 120_000 });
  await expect(updatedFragment.locator('math')).toHaveCount(0);
  expect(errors).toEqual([]);
});
