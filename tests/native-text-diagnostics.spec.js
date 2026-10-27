/** @jest-environment jsdom */
import { recordNativeTextMismatch, getNativeTextDiagnostics } from '../src/lib/native-text-diagnostics.js';

test('native mismatch diagnostics retain bounded text evidence and exclude binary package parts', () => {
    const xml = '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
        + '<pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Native sample</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part>'
        + '<pkg:part pkg:name="/word/media/image1.png"><pkg:binaryData>DO_NOT_EXPORT_BINARY</pkg:binaryData></pkg:part></pkg:package>';
    for (let i = 0; i < 12; i++) recordNativeTextMismatch({ value: xml, currentText: { value: `native ${i}` } }, { text: `expected ${i}` });
    const records = getNativeTextDiagnostics();
    expect(records).toHaveLength(8);
    expect(records[0]).toMatchObject({ expectedText: 'expected 4', reviewedText: 'native 4' });
    expect(records[0].documentXml).toContain('Native sample');
    expect(JSON.stringify(records)).not.toContain('DO_NOT_EXPORT_BINARY');
    records[0].expectedText = 'mutated caller copy';
    expect(getNativeTextDiagnostics()[0].expectedText).toBe('expected 4');
});
