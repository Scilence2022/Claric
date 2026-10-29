/** Resolve the picture's own section, never the current caret's section. */
export async function readImageTextWidth(context, pic) {
    try {
        const setup = pic.getRange().sections.getFirst().pageSetup;
        const paragraph = pic.paragraph;
        const table = pic.parentTableOrNullObject;
        setup.load('pageWidth,leftMargin,rightMargin,gutter');
        setup.textColumns.load('items');
        paragraph.load('leftIndent,rightIndent,firstLineIndent');
        table.load('isNullObject');
        await context.sync();
        if (!table.isNullObject) throw new Error('Images in tables require an explicit cell-sized width.');
        if (setup.textColumns.items.length !== 1) throw new Error('Multi-column sections require an explicit column-sized width.');
        let gutter = setup.gutter;
        if (gutter > 0) {
            setup.load('gutterPosition');
            await context.sync();
            if (String(setup.gutterPosition).toLowerCase() === 'top') gutter = 0;
        }
        const values = [setup.pageWidth, setup.leftMargin, setup.rightMargin, gutter,
            paragraph.leftIndent, paragraph.rightIndent, paragraph.firstLineIndent];
        if (!values.every(Number.isFinite)) throw new Error('Incomplete page or paragraph dimensions.');
        const width = setup.pageWidth - setup.leftMargin - setup.rightMargin - gutter
            - paragraph.leftIndent - paragraph.rightIndent - Math.max(0, paragraph.firstLineIndent);
        if (width < 8 || width > 1584) throw new Error('Invalid available text width.');
        return width;
    } catch (error) {
        throw new Error(`Cannot fit image to text width: ${error.message}. Readable section layout is required; no fixed width was substituted.`);
    }
}

