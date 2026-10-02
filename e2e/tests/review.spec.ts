/**
 * Word-style review features
 *
 * Ribbon + modes (Editing / Reviewing / Viewing), comments in margin balloons,
 * Track Changes (suggestions), markup views, Word tags (content controls),
 * page setup, status bar, File backstage and print.
 */

import { test, expect, type Page } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { EditorPage } from '../helpers/editor-page';
import * as assertions from '../helpers/assertions';

/** Save through the editor ref and return the DOCX as base64 */
async function saveBase64(page: Page): Promise<string> {
  return page.evaluate(async () => {
    const ref = (
      window as unknown as {
        __docxEditorRef: { current: { save: () => Promise<ArrayBuffer | null> } };
      }
    ).__docxEditorRef;
    const buffer = await ref.current.save();
    if (!buffer) return '';
    let binary = '';
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
    return btoa(binary);
  });
}

/** Save through the editor ref and return the DOCX parts as text */
async function saveAndUnzip(page: Page): Promise<Record<string, string>> {
  const base64 = await saveBase64(page);
  expect(base64).not.toBe('');
  const zip = await JSZip.loadAsync(Buffer.from(base64, 'base64'));
  const files: Record<string, string> = {};
  for (const [path, file] of Object.entries(zip.files)) {
    if (!file.dir && /\.(xml|rels)$/.test(path)) files[path] = await file.async('text');
  }
  return files;
}

async function setMode(page: Page, mode: 'editing' | 'suggesting' | 'viewing') {
  await page.getByTestId('mode-menu').click();
  await page.getByTestId(`mode-${mode}`).click();
}

async function openReviewTab(page: Page) {
  await page.getByTestId('ribbon-tab-review').click();
}

async function addComment(page: Page, editor: EditorPage, anchor: string, text: string) {
  await editor.selectText(anchor);
  await openReviewTab(page);
  await page.getByTestId('review-add-comment').click();
  const composer = page.getByTestId('comment-composer');
  await composer.locator('textarea').fill(text);
  await composer.getByRole('button', { name: 'Comment' }).click();
}

async function setup(page: Page, text: string): Promise<EditorPage> {
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.focus();
  await editor.typeText(text);
  return editor;
}

test.describe('Ribbon and modes', () => {
  test('ribbon shows Word tabs and starts in Editing mode', async ({ page }) => {
    await setup(page, 'Hello');
    for (const tab of ['home', 'insert', 'layout', 'review', 'view', 'developer']) {
      await expect(page.getByTestId(`ribbon-tab-${tab}`)).toBeVisible();
    }
    await expect(page.getByTestId('ribbon-tab-home')).toHaveAttribute('aria-selected', 'true');
    await expect(page.getByTestId('mode-menu')).toContainText('Editing');
  });

  test('arrow keys move between ribbon tabs', async ({ page }) => {
    await setup(page, 'Hello');
    await page.getByTestId('ribbon-tab-home').focus();
    await page.keyboard.press('ArrowRight');
    await expect(page.getByTestId('ribbon-tab-insert')).toHaveAttribute('aria-selected', 'true');
    await expect(page.getByTestId('ribbon-tab-insert')).toBeFocused();
  });

  test('Viewing mode is read-only and hides editing tabs', async ({ page }) => {
    const editor = await setup(page, 'Locked text');
    await setMode(page, 'viewing');

    await expect(page.getByTestId('ribbon-tab-home')).toHaveCount(0);
    await expect(page.locator('[contenteditable="true"]')).toHaveCount(0);
    await page.locator('.layout-page').first().click();
    await page.keyboard.type('XYZ');
    await expect(page.locator('.layout-page')).not.toContainText('XYZ');
    await expect(page.getByTestId('status-bar')).toContainText('Read-only');

    await setMode(page, 'editing');
    await expect(page.locator('[data-testid="toolbar-bold"]')).toHaveCount(1);
    await editor.focus();
    await expect(page.locator('[contenteditable="true"]').first()).toBeAttached();
  });

  test('Home stays mounted: Ctrl+B works while another tab is shown', async ({ page }) => {
    const editor = await setup(page, 'Bold me');
    await page.getByTestId('ribbon-tab-view').click();
    await editor.selectText('Bold');
    await editor.applyBoldShortcut();
    await page.getByTestId('ribbon-tab-home').click();
    await expect(page.getByTestId('toolbar-bold')).toHaveAttribute('aria-pressed', 'true');
  });
});

test.describe('Track Changes (suggestions)', () => {
  test('Ctrl+Shift+E and the status bar toggle Track Changes', async ({ page }) => {
    const editor = await setup(page, 'Hello');
    const status = page.getByTestId('status-track-changes');
    await expect(status).toContainText('Off');

    await editor.focus();
    await page.keyboard.press('Control+Shift+E');
    await expect(status).toContainText('On');
    await expect(page.getByTestId('mode-menu')).toContainText('Reviewing');

    await status.click();
    await expect(status).toContainText('Off');
  });

  test('typed text becomes an insertion; Accept on the ribbon applies it', async ({ page }) => {
    const editor = await setup(page, 'Hello world');
    await openReviewTab(page);
    await page.getByTestId('ribbon-track-changes').click();

    await editor.focus();
    await page.keyboard.press('End');
    await page.keyboard.type(' again');
    const ins = page.locator('.layout-page .docx-revision-ins');
    await expect(ins.first()).toBeVisible();
    await expect(ins).toContainText(['again']);

    // Cursor sits inside the change: Accept accepts it
    await page.keyboard.press('ArrowLeft');
    await page.getByTestId('ribbon-accept').click();
    await expect(page.locator('.layout-page .docx-revision-ins')).toHaveCount(0);
    await assertions.assertDocumentContainsText(page, 'Hello world again');
  });

  test('deleted text is struck through and can be rejected from the Reviewing Pane', async ({
    page,
  }) => {
    const editor = await setup(page, 'Hello world');
    await setMode(page, 'suggesting');
    await editor.selectText('world');
    await page.keyboard.press('Backspace');

    const del = page.locator('.layout-page .docx-revision-del');
    await expect(del.first()).toHaveText('world');

    await openReviewTab(page);
    await page.getByTestId('review-toggle-sidebar').click();
    await page.getByTestId('review-tab-suggestions').click();
    await page.getByTestId('reject-suggestion').click();
    await expect(page.locator('.layout-page .docx-revision-del')).toHaveCount(0);
    await assertions.assertDocumentContainsText(page, 'Hello world');
  });

  test('Accept All Changes applies insertions and deletions', async ({ page }) => {
    const editor = await setup(page, 'Hello world');
    await setMode(page, 'suggesting');
    await editor.selectText('world');
    await page.keyboard.type('there');

    await openReviewTab(page);
    await page.getByTestId('ribbon-accept-menu').click();
    await page.getByTestId('accept-all-changes').click();

    await expect(page.locator('.layout-page .docx-revision-ins')).toHaveCount(0);
    await expect(page.locator('.layout-page .docx-revision-del')).toHaveCount(0);
    await assertions.assertDocumentContainsText(page, 'Hello there');
    await expect(page.locator('.layout-page')).not.toContainText('world');
  });

  test('markup views: No Markup shows final text, Original shows the original', async ({
    page,
  }) => {
    const editor = await setup(page, 'Hello world');
    await setMode(page, 'suggesting');
    await editor.selectText('world');
    await page.keyboard.type('there');
    await openReviewTab(page);

    await page.getByTestId('ribbon-markup-view').click();
    await page.getByTestId('markup-none').click();
    await expect(page.locator('.layout-page .docx-revision-del').first()).toBeHidden();
    await expect(page.locator('.layout-page .docx-revision-ins').first()).toBeVisible();

    await page.getByTestId('ribbon-markup-view').click();
    await page.getByTestId('markup-original').click();
    await expect(page.locator('.layout-page .docx-revision-ins').first()).toBeHidden();
    await expect(page.locator('.layout-page .docx-revision-del').first()).toBeVisible();
  });

  test('suggestions save as w:ins / w:del with w:delText', async ({ page }) => {
    const editor = await setup(page, 'Hello world');
    await setMode(page, 'suggesting');
    await editor.selectText('world');
    await page.keyboard.type('there');

    const xml = (await saveAndUnzip(page))['word/document.xml'];
    expect(xml).toMatch(/<w:ins [^>]*w:author="Demo User"/);
    expect(xml).toMatch(
      /<w:del [^>]*w:author="Demo User"[^>]*>.*<w:delText[^>]*>world<\/w:delText>/
    );
    expect(xml).toMatch(/<w:ins [^>]*>.*<w:t[^>]*>there<\/w:t>/);
  });
});

test.describe('Comments (margin balloons)', () => {
  test('New Comment opens a balloon next to the text', async ({ page }) => {
    const editor = await setup(page, 'The quarterly report is ready');
    await addComment(page, editor, 'quarterly report', 'Please double-check the figures');

    const thread = page.getByTestId('comment-thread');
    await expect(thread).toHaveCount(1);
    await expect(thread).toContainText('Please double-check the figures');
    await expect(thread).toContainText('Demo User');
    await expect(page.locator('.layout-page .docx-comment-anchor').first()).toBeVisible();

    // Balloon sits in the margin to the right of the page
    const pageBox = await page.locator('.layout-page').first().boundingBox();
    const balloonBox = await thread.boundingBox();
    expect(balloonBox!.x).toBeGreaterThan(pageBox!.x + pageBox!.width);
  });

  test('Ctrl+Alt+M adds a comment', async ({ page }) => {
    const editor = await setup(page, 'Shortcut comment');
    await editor.selectText('Shortcut');
    await page.keyboard.press('Control+Alt+M');
    await expect(page.getByTestId('comment-composer')).toBeVisible();
    await expect(page.getByTestId('comment-composer').locator('textarea')).toBeFocused();
  });

  test('selecting a word enables New Comment', async ({ page }) => {
    const editor = await setup(page, 'The quarterly report is ready');
    await editor.selectText('quarterly');
    await page.getByTestId('ribbon-tab-review').click();
    await expect(page.getByTestId('review-add-comment')).toBeEnabled();
    await page.getByTestId('review-add-comment').click();
    await expect(page.getByTestId('comment-composer')).toBeVisible();
    await page.getByTestId('comment-composer').locator('textarea').fill('Looks good');
    await page.getByTestId('comment-composer').getByRole('button', { name: 'Comment' }).click();
    await expect(page.getByTestId('comment-thread')).toContainText('Looks good');
    await expect(page.locator('.layout-page .docx-comment-anchor').first()).toBeVisible();
  });

  test('selection shows a comment icon at the right edge of the page', async ({ page }) => {
    const editor = await setup(page, 'The quarterly report is ready');
    await editor.selectText('quarterly');
    const btn = page.getByTestId('selection-comment-button');
    await expect(btn).toBeVisible({ timeout: 5000 });

    // Icon is pinned to the right end of the page (not on top of the highlight)
    const pageBox = await page.locator('.layout-page').first().boundingBox();
    const sel = page.getByTestId('selection-rect-0');
    await expect(sel).toBeVisible();
    const selBox = await sel.boundingBox();
    const btnBox = await btn.boundingBox();
    expect(pageBox).toBeTruthy();
    expect(selBox).toBeTruthy();
    expect(btnBox).toBeTruthy();
    expect(btnBox!.x).toBeGreaterThan(selBox!.x + selBox!.width);
    // Within ~40px of the page's right edge
    expect(btnBox!.x + btnBox!.width).toBeGreaterThan(pageBox!.x + pageBox!.width - 40);

    await btn.click();
    await expect(page.getByTestId('comment-composer')).toBeVisible();
    // Icon hides while the draft composer is open
    await expect(page.getByTestId('selection-comment-button')).toHaveCount(0);
    await page.getByTestId('comment-composer').locator('textarea').fill('From selection icon');
    await page.getByTestId('comment-composer').getByRole('button', { name: 'Comment' }).click();
    await expect(page.getByTestId('comment-thread')).toContainText('From selection icon');
  });

  test('reply to and resolve a comment', async ({ page }) => {
    const editor = await setup(page, 'The quarterly report is ready');
    await addComment(page, editor, 'quarterly report', 'First note');

    await page.getByTestId('comment-reply').click();
    const reply = page.getByTestId('reply-composer');
    await reply.locator('textarea').fill('Done, figures verified');
    await reply.getByRole('button', { name: 'Reply' }).click();
    await expect(page.getByTestId('comment-thread')).toContainText('Done, figures verified');

    await page.getByTestId('comment-resolve').click();
    await expect(page.getByTestId('comment-thread')).toContainText('Resolved');
    await expect(page.getByTestId('comment-thread')).toHaveClass(/is-resolved/);
  });

  test('Simple Markup collapses balloons to icons', async ({ page }) => {
    const editor = await setup(page, 'The quarterly report is ready');
    await addComment(page, editor, 'quarterly', 'Note');
    await editor.focus();
    await page.keyboard.press('End');

    await page.getByTestId('ribbon-markup-view').click();
    await page.getByTestId('markup-simple').click();
    await expect(page.getByTestId('comment-indicator')).toHaveCount(1);
    await page.getByTestId('comment-indicator').click();
    await expect(page.getByTestId('comment-thread')).toContainText('Note');
  });

  test('Previous/Next and Delete on the ribbon', async ({ page }) => {
    const editor = await setup(page, 'Alpha beta gamma');
    await addComment(page, editor, 'Alpha', 'One');
    await addComment(page, editor, 'gamma', 'Two');
    await expect(page.getByTestId('comment-thread')).toHaveCount(2);

    await page.getByTestId('ribbon-prev-comment').click();
    await page.getByTestId('ribbon-delete-comment').click();
    await expect(page.getByTestId('comment-thread')).toHaveCount(1);
    await expect(page.getByTestId('comment-thread')).toContainText('Two');
  });

  test('delete thread from the balloon menu', async ({ page }) => {
    const editor = await setup(page, 'The quarterly report is ready');
    await addComment(page, editor, 'quarterly', 'Temporary');
    await page.getByRole('button', { name: 'More thread actions' }).click();
    await page.getByTestId('comment-delete').click();
    await expect(page.getByTestId('comment-thread')).toHaveCount(0);
    await expect(page.locator('.layout-page .docx-comment-anchor')).toHaveCount(0);
  });

  test('comments save to comments.xml with references and replies', async ({ page }) => {
    const editor = await setup(page, 'The quarterly report is ready');
    await addComment(page, editor, 'quarterly report', 'Check this');
    await page.getByTestId('comment-reply').click();
    await page.getByTestId('reply-composer').locator('textarea').fill('Checked');
    await page.getByTestId('reply-composer').getByRole('button', { name: 'Reply' }).click();

    const files = await saveAndUnzip(page);
    const doc = files['word/document.xml'];
    expect(doc).toContain('<w:commentRangeStart w:id="0"/>');
    expect(doc).toContain('<w:commentRangeEnd w:id="0"/>');
    expect(doc).toContain('<w:commentReference w:id="0"/>');
    expect(files['word/comments.xml']).toMatch(/<w:comment w:id="0" w:author="Demo User"/);
    expect(files['word/comments.xml']).toContain('Check this');
    expect(files['word/comments.xml']).toContain('Checked');
    expect(files['word/commentsExtended.xml']).toContain('w15:paraIdParent');
    expect(files['word/_rels/document.xml.rels']).toContain('Target="comments.xml"');
    expect(files['[Content_Types].xml']).toContain('PartName="/word/comments.xml"');
  });
});

test.describe('Word tags (Developer > Controls)', () => {
  test('wrap selected text in a content control, set Title/Tag, save as w:sdt', async ({
    page,
  }) => {
    const editor = await setup(page, 'Dear Jane Doe,');
    await editor.selectText('Jane Doe');
    await page.getByTestId('ribbon-tab-developer').click();
    await page.getByTestId('review-insert-tag').click();
    await expect(page.locator('.layout-page .docx-content-control').first()).toBeVisible();

    await page.getByTestId('ribbon-control-properties').click();
    const dialog = page.getByTestId('control-properties-dialog');
    await dialog.getByTestId('tag-title-input').fill('Client name');
    await dialog.getByTestId('tag-tag-input').fill('client_name');
    await dialog.getByRole('button', { name: 'OK' }).click();
    await expect(dialog).toHaveCount(0);

    await page.getByTestId('ribbon-tags-pane').click();
    await expect(page.getByTestId('tag-card')).toContainText('client_name');

    const xml = (await saveAndUnzip(page))['word/document.xml'];
    expect(xml).toMatch(
      /<w:sdt><w:sdtPr><w:alias w:val="Client name"\/><w:tag w:val="client_name"\/>.*Jane Doe.*<\/w:sdt>/
    );
  });

  test('lock content control on create/edit and persist w:lock', async ({ page }) => {
    const editor = await setup(page, 'Dear Jane Doe,');
    await editor.selectText('Jane Doe');
    await page.getByTestId('ribbon-tab-developer').click();
    await page.getByTestId('review-insert-tag').click();
    await expect(page.locator('.layout-page .docx-content-control').first()).toBeVisible();

    await page.getByTestId('ribbon-control-properties').click();
    const dialog = page.getByTestId('control-properties-dialog');
    await dialog.getByTestId('tag-title-input').fill('Client name');
    await dialog.getByTestId('tag-tag-input').fill('client_name');
    await dialog.getByTestId('tag-lock-cannot-delete').check();
    await dialog.getByTestId('tag-lock-cannot-edit').check();
    await dialog.getByRole('button', { name: 'OK' }).click();
    await expect(dialog).toHaveCount(0);

    await expect(page.locator('.layout-page .docx-content-control--locked').first()).toBeVisible();
    await page.getByTestId('ribbon-tags-pane').click();
    await expect(page.getByTestId('tag-lock-badge')).toBeVisible();
    await expect(page.getByTestId('tag-card').getByText('Remove tag')).toHaveCount(0);

    const xml = (await saveAndUnzip(page))['word/document.xml'];
    expect(xml).toContain('<w:lock w:val="sdtContentLocked"/>');
    expect(xml).toMatch(/<w:alias w:val="Client name"\/>/);
    expect(xml).toMatch(/<w:tag w:val="client_name"\/>/);
  });
});

test.describe('Layout, Insert and status bar', () => {
  test('Orientation > Landscape makes pages wider than tall', async ({ page }) => {
    await setup(page, 'Landscape');
    await page.getByTestId('ribbon-tab-layout').click();
    await page.getByTestId('ribbon-orientation').click();
    await page.getByTestId('orientation-landscape').click();
    await expect
      .poll(async () => {
        const box = await page.locator('.layout-page').first().boundingBox();
        return box ? box.width > box.height : false;
      })
      .toBe(true);
  });

  test('Insert > Page Break starts a new page', async ({ page }) => {
    const editor = await setup(page, 'First page');
    await editor.focus();
    await page.keyboard.press('End');
    await page.getByTestId('ribbon-tab-insert').click();
    await page.getByTestId('ribbon-page-break').click();
    await page.keyboard.type('Second page');
    await expect(page.locator('.layout-page')).toHaveCount(2);
    await expect(page.getByTestId('status-page')).toContainText('of 2');

    const xml = (await saveAndUnzip(page))['word/document.xml'];
    expect(xml).toContain('<w:pageBreakBefore/>');
  });

  test('status bar shows word count and opens Word Count', async ({ page }) => {
    await setup(page, 'one two three four');
    await expect(page.getByTestId('status-words')).toContainText('4 words');
    await page.getByTestId('status-words').click();
    await expect(page.getByTestId('word-count-words')).toHaveText('4');
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('word-count-dialog')).toHaveCount(0);
  });

  test('status bar zoom changes the zoom level', async ({ page }) => {
    await setup(page, 'Zoom');
    await page.getByTestId('status-bar').getByRole('button', { name: 'Zoom in' }).click();
    await expect(page.getByTestId('status-zoom')).toHaveText('110%');
  });
});

test.describe('File backstage and print', () => {
  test('File opens Info; Escape returns to the document', async ({ page }) => {
    await setup(page, 'Some words here');
    await page.getByTestId('ribbon-tab-file').click();
    await expect(page.getByTestId('backstage')).toBeVisible();
    await expect(page.getByTestId('backstage')).toContainText('Words');
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('backstage')).toHaveCount(0);
  });

  test('File > Print writes pages and markup styles to the print window', async ({ page }) => {
    await setup(page, 'Printable content');
    await page.evaluate(() => {
      const captured: string[] = [];
      (window as unknown as { __printed: string[] }).__printed = captured;
      window.open = (() => ({
        document: { write: (html: string) => captured.push(html), close: () => {} },
        print: () => {},
        close: () => {},
        closed: true,
        set onload(_fn: unknown) {},
      })) as unknown as typeof window.open;
    });

    await page.getByTestId('ribbon-tab-file').click();
    await page.getByTestId('backstage-print').click();
    await page.getByTestId('backstage-print-button').click();
    const html = await page.evaluate(() =>
      (window as unknown as { __printed: string[] }).__printed.join('')
    );
    expect(html).toContain('Printable content');
    expect(html).toContain('.docx-revision-');
    expect(html).toContain('layout-page');
  });
});

test.describe('Round trip', () => {
  test('comments, replies, resolved state and suggestions survive save + reopen', async ({
    page,
  }, testInfo) => {
    const editor = await setup(page, 'Alpha beta gamma');
    await addComment(page, editor, 'beta', 'Is beta right?');
    await page.getByTestId('comment-reply').click();
    await page.getByTestId('reply-composer').locator('textarea').fill('Yes');
    await page.getByTestId('reply-composer').getByRole('button', { name: 'Reply' }).click();
    await page.getByTestId('comment-resolve').click();

    await setMode(page, 'suggesting');
    await editor.selectText('gamma');
    await page.keyboard.press('Backspace');
    await expect(page.locator('.layout-page .docx-revision-del').first()).toBeVisible();

    const file = testInfo.outputPath('roundtrip.docx');
    await writeFile(file, Buffer.from(await saveBase64(page), 'base64'));

    await page.reload();
    await editor.waitForReady();
    await editor.loadDocxFile(file);

    await expect(page.locator('.layout-page .docx-revision-del').first()).toHaveText('gamma');
    const thread = page.getByTestId('comment-thread');
    await expect(thread).toHaveCount(1);
    await expect(thread).toContainText('Is beta right?');
    await expect(thread).toContainText('Yes');
    await expect(thread).toContainText('Resolved');

    await openReviewTab(page);
    await page.getByTestId('review-toggle-sidebar').click();
    await page.getByTestId('review-tab-suggestions').click();
    await expect(page.getByTestId('suggestion-card')).toContainText('gamma');
  });
});

test.describe('Header & Footer (Insert tab)', () => {
  test('Footer > Page X of Y shows page numbers and saves a footer part', async ({ page }) => {
    await setup(page, 'Body text');
    await page.getByTestId('ribbon-tab-insert').click();
    await page.getByTestId('ribbon-footer').click();
    await page.getByTestId('footer-page-x-of-y').click();

    await expect(page.locator('.layout-page-footer .layout-hf-content').first()).toHaveText(
      'Page 1 of 1'
    );

    const files = await saveAndUnzip(page);
    const footerPath = Object.keys(files).find((p) => /^word\/footer\d+\.xml$/.test(p));
    expect(footerPath).toBeTruthy();
    const footer = files[footerPath!];
    expect(footer).toContain('<w:ftr ');
    expect(footer).toMatch(/w:instr=" PAGE "/);
    expect(footer).toMatch(/w:instr=" NUMPAGES "/);
    expect(files['word/document.xml']).toMatch(
      /<w:footerReference w:type="default" r:id="rIdDecfooter1"\/>/
    );
    expect(files['word/_rels/document.xml.rels']).toContain('Id="rIdDecfooter1"');
    expect(files['[Content_Types].xml']).toContain(`PartName="/${footerPath}"`);
  });

  test('Header > Blank opens in-document editor; text is saved', async ({ page }) => {
    await setup(page, 'Body text');
    await page.getByTestId('ribbon-tab-insert').click();
    await page.getByTestId('ribbon-header').click();
    await page.getByTestId('header-blank').click();

    const editor = page.getByTestId('header-footer-editor');
    await expect(editor).toBeVisible();
    // Editor is positioned on the page (not a centered modal popup)
    await expect(page.locator('.paged-editor__pages.is-hf-editing-header')).toBeVisible();
    await editor.locator('.hf-editor-pm [contenteditable="true"]').click();
    await page.keyboard.type('Confidential');
    await page.getByTestId('hf-editor-close').click();

    await expect(page.locator('.layout-page-header').first()).toContainText('Confidential');
    const files = await saveAndUnzip(page);
    const headerPath = Object.keys(files).find((p) => /^word\/header\d+\.xml$/.test(p));
    expect(files[headerPath!]).toContain('Confidential');
  });

  test('header editor uses ribbon formatting and can insert page number', async ({ page }) => {
    await setup(page, 'Body text');
    await page.getByTestId('ribbon-tab-insert').click();
    await page.getByTestId('ribbon-header').click();
    await page.getByTestId('header-blank').click();

    const hf = page.getByTestId('header-footer-editor');
    await expect(hf).toBeVisible();
    const pm = hf.locator('.hf-editor-pm [contenteditable="true"]');
    await pm.click();
    await page.keyboard.type('Draft');

    // Insert → Page Number at caret (collapsed) while editing header
    await page.getByTestId('ribbon-page-number').click();
    await page.getByTestId('page-number-current').click();
    await expect(hf.locator('.docx-field, [data-field-type="PAGE"]').first()).toBeVisible();
    await expect(hf.locator('.hf-editor-pm')).toContainText('Draft');

    // Bold the word via Home ribbon — proves formatting targets the HF view
    await pm.click();
    const mod = process.platform === 'darwin' ? 'Meta' : 'Control';
    await page.keyboard.press(`${mod}+ArrowLeft`); // start of "Draft" or line
    await page.keyboard.press(`Shift+${mod}+ArrowRight`);
    await page.getByTestId('ribbon-tab-home').click();
    await page.getByTestId('toolbar-bold').click();
    await expect(hf.locator('strong').first()).toBeVisible();

    await page.getByTestId('hf-editor-close').click();
    const header = page.locator('.layout-page-header .layout-hf-content').first();
    await expect(header).toContainText('Draft');
    await expect(header).toContainText('1');
  });

  test('double-click header zone opens in-document editor', async ({ page }) => {
    await setup(page, 'Body text');
    // Empty header zone is always painted for Word-style hover / double-click
    const zone = page.locator('.layout-page-header').first();
    await expect(zone).toBeVisible();
    await zone.dblclick();
    await expect(page.getByTestId('header-footer-editor')).toBeVisible();
    await expect(page.locator('.paged-editor__pages.is-hf-editing-header')).toBeVisible();
  });

  test('Remove Footer clears footer content from the page', async ({ page }) => {
    await setup(page, 'Body text');
    await page.getByTestId('ribbon-tab-insert').click();
    await page.getByTestId('ribbon-footer').click();
    await page.getByTestId('footer-page-number').click();
    await expect(page.locator('.layout-page-footer .layout-hf-content').first()).toHaveText('1');

    await page.getByTestId('ribbon-footer').click();
    await page.getByTestId('footer-remove').click();
    // Hit zone remains (for hover/edit) but is empty
    await expect(page.locator('.layout-page-footer').first()).toHaveClass(/is-empty/);
    await expect(page.locator('.layout-page-footer .layout-hf-content').first()).toHaveText('');
  });
});

test.describe('Ribbon layout', () => {
  test('ribbon is thin by default and can switch to Classic', async ({ page }) => {
    await setup(page, 'Hello');
    const ribbon = page.getByTestId('ribbon');
    await expect(ribbon).toHaveClass(/ep-ribbon--simplified/);
    const thin = (await ribbon.boundingBox())!.height;
    expect(thin).toBeLessThan(90);

    await page.getByTestId('ribbon-display-options').click();
    await page.getByTestId('ribbon-layout-classic').click();
    await expect(ribbon).toHaveClass(/ep-ribbon--classic/);
    expect((await ribbon.boundingBox())!.height).toBeGreaterThan(thin);

    await page.getByTestId('ribbon-display-options').click();
    await page.getByTestId('ribbon-layout-tabs-only').click();
    await expect(page.locator('#ep-ribbon-panel-home')).toBeHidden();
  });
});

test.describe('File: New, Open, Save', () => {
  test('New asks before discarding changes, then starts a blank document', async ({ page }) => {
    await setup(page, 'Unsaved work');
    await page.getByTestId('ribbon-tab-file').click();
    await page.getByTestId('backstage-new').click();
    await page.getByTestId('backstage-new-blank').click();

    const dialog = page.getByTestId('discard-dialog');
    await expect(dialog).toBeVisible();
    await dialog.getByRole('button', { name: 'Discard' }).click();

    await expect(page.locator('.layout-page')).not.toContainText('Unsaved work');
    await expect(page.getByTestId('status-words')).toContainText('0 words');
  });

  test('Open loads a .docx from the file picker', async ({ page }) => {
    await setup(page, '');
    await page.getByTestId('ribbon-tab-file').click();
    await page.getByTestId('backstage-open').click();
    const chooser = page.waitForEvent('filechooser');
    await page.getByTestId('backstage-browse').click();
    await (await chooser).setFiles('e2e/fixtures/styled-content.docx');
    await expect(page.locator('.layout-page').first()).not.toBeEmpty();
    await page.getByTestId('ribbon-tab-file').click();
    await expect(page.getByTestId('backstage')).toContainText('styled-content.docx');
  });

  test('Save downloads the .docx when the host has no onSave; Ctrl+S too', async ({ page }) => {
    const editor = await setup(page, 'Save me');
    await page.getByTestId('ribbon-tab-file').click();
    const download = page.waitForEvent('download');
    await page.getByTestId('backstage-save').click();
    expect((await download).suggestedFilename()).toMatch(/\.docx$/);

    await editor.focus();
    const download2 = page.waitForEvent('download');
    await page.keyboard.press('Control+S');
    expect((await download2).suggestedFilename()).toMatch(/\.docx$/);
  });
});

test.describe('Alignment with side panes open', () => {
  test('ruler and Word tags pane stay aligned when the review sidebar opens', async ({ page }) => {
    // Wide desktop window: page + tags pane fit side by side
    await page.setViewportSize({ width: 1600, height: 900 });
    const editor = await setup(page, 'Dear Jane Doe,');
    await editor.selectText('Jane Doe');
    await page.getByTestId('ribbon-tab-developer').click();
    await page.getByTestId('review-insert-tag').click();
    await expect(page.locator('.layout-page .docx-content-control').first()).toBeVisible();

    await page.getByTestId('ribbon-tags-pane').click();
    await expect(page.getByTestId('tag-card').first()).toBeVisible();

    await expect
      .poll(() =>
        page.evaluate(() => {
          const pg = document.querySelector('.layout-page')!.getBoundingClientRect();
          const ruler = document.querySelector('.ep-ruler-row > *')!.getBoundingClientRect();
          const pane = document
            .querySelector('[data-testid="review-sidebar"]')!
            .getBoundingClientRect();
          return {
            rulerDx: Math.round(ruler.left - pg.left),
            panelOnScreen: pane.right <= window.innerWidth,
          };
        })
      )
      .toEqual({ rulerDx: 0, panelOnScreen: true });
  });
});
