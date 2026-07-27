# Word object-model bridge

Use `word_com` when a focused World tool cannot express an operation.

## Roots

- `application`: the running `Word.Application`
- `document`: the selected open document or `ActiveDocument`
- `selection`: `Application.Selection`
- `handle`: a COM object returned by an earlier `word_com` call

## Operations

- `get`: read a property path, such as `PageSetup.Orientation`.
- `set`: set the final property in a path.
- `call`: call the final method in a path with positional `arguments`.
- `release`: release one retained handle.
- `release_all`: release all retained handles.

Returned COM objects include a `handle`. Use that handle as the next root to traverse collections or nested objects. For example:

1. Get `Sections` from the document and retain its handle.
2. Call `Item` on that handle with argument `1`.
3. Get `Headers` from the returned section handle.
4. Call `Item` with the requested Word constant.
5. Set `Range.Text` on the returned header handle.

Word COM constants are integers. Common values:

- `wdCollapseStart = 1`, `wdCollapseEnd = 0`
- `wdReplaceNone = 0`, `wdReplaceOne = 1`, `wdReplaceAll = 2`
- `wdFindStop = 0`, `wdFindContinue = 1`
- `wdAlignParagraphLeft = 0`, center `1`, right `2`, justify `3`
- `wdHeaderFooterPrimary = 1`, first page `2`, even pages `3`
- `wdExportFormatPDF = 17`
- `wdDoNotSaveChanges = 0`, `wdSaveChanges = -1`

Treat macro execution, printing, protection changes, mail merge sends, and destructive bulk operations as explicit-only actions.
