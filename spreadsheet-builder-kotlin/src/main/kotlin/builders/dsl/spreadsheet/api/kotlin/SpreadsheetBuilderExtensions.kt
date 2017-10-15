package builders.dsl.spreadsheet.api.kotlin

import builders.dsl.spreadsheet.api.*
import builders.dsl.spreadsheet.builder.api.*
import builders.dsl.spreadsheet.query.api.*


fun CanDefineStyle.style(name: String, styleDefinition: CellStyleDefinition.(CellStyleDefinition) -> CellStyleDefinition): CanDefineStyle {
    return this.style(name) {
        styleDefinition(it, it)
    }
}

fun CellDefinition.comment(commentDefinition: CommentDefinition.(CommentDefinition) -> CommentDefinition): CellDefinition {
    return this.comment {
        commentDefinition(it, it)
    }
}

fun CellDefinition.text(text: String, fontConfiguration: FontDefinition.(FontDefinition) -> FontDefinition): CellDefinition {
    return this.text(text) {
        fontConfiguration(it, it)
    }
}

fun CellStyleDefinition.font(fontConfiguration: FontDefinition.(FontDefinition) -> FontDefinition): CellStyleDefinition {
    return this.font {
        fontConfiguration(it, it)
    }
}

/**
 * Configures all the borders of the cell.
 * @param borderConfiguration border configuration closure
 */
fun CellStyleDefinition.border(borderConfiguration: BorderDefinition.(BorderDefinition) -> BorderDefinition): CellStyleDefinition {
    return this.border {
        borderConfiguration(it, it)
    }
}

/**
 * Configures one border of the cell.
 * @param location border to be configured
 * @param borderConfiguration border configuration closure
 */
fun CellStyleDefinition.border(location: Keywords.BorderSide, borderConfiguration: BorderDefinition.(BorderDefinition) -> BorderDefinition): CellStyleDefinition {
    return this.border(location) {
        borderConfiguration(it, it);
    }
}

/**
 * Configures two borders of the cell.
 * @param first first border to be configured
 * @param second second border to be configured
 * @param borderConfiguration border configuration closure
 */
fun CellStyleDefinition.border(first: Keywords.BorderSide, second: Keywords.BorderSide, borderConfiguration: BorderDefinition.(BorderDefinition) -> BorderDefinition): CellStyleDefinition {
    return this.border(first, second) {
        borderConfiguration(it, it)
    }
}

/**
 * Configures three borders of the cell.
 * @param first first border to be configured
 * @param second second border to be configured
 * @param third third border to be configured
 * @param borderConfiguration border configuration closure
 */
fun CellStyleDefinition.border(first: Keywords.BorderSide, second: Keywords.BorderSide, third: Keywords.BorderSide, borderConfiguration: BorderDefinition.(BorderDefinition) -> BorderDefinition): CellStyleDefinition {
    return this.border(first, second, third) {
        borderConfiguration(it, it)
    }
}

/**
 * Applies a customized named style to the current element.
 *
 * @param name the name of the style
 * @param styleDefinition the definition of the style customizing the predefined style
 */
fun HasStyle.style(name: String, styleDefinition: CellStyleDefinition.(CellStyleDefinition) -> CellStyleDefinition): HasStyle {
    return this.style(name) {
        styleDefinition(it, it)
    }
}

/**
 * Applies a customized named style to the current element.
 *
 * @param names the names of the styles
 * @param styleDefinition the definition of the style customizing the predefined style
 */
fun HasStyle.styles(names: Iterable<String>, styleDefinition: CellStyleDefinition.(CellStyleDefinition) -> CellStyleDefinition): HasStyle {
    return this.styles(names) {
        styleDefinition(it, it)
    }
}

/**
 * Applies the style defined by the closure to the current element.
 * @param styleDefinition the definition of the style
 */
fun HasStyle.style(styleDefinition: CellStyleDefinition.(CellStyleDefinition) -> CellStyleDefinition): HasStyle {
    return this.style {
        styleDefinition(it, it)
    }
}

fun RowDefinition.cell(cellDefinition: CellDefinition.(CellDefinition) -> CellDefinition): RowDefinition {
    return this.cell {
        cellDefinition(it, it)
    }
}

fun RowDefinition.cell(column: Int, cellDefinition: CellDefinition.(CellDefinition) -> CellDefinition): RowDefinition {
    return this.cell(column) {
        cellDefinition(it, it)
    }
}

fun RowDefinition.cell(column: String, cellDefinition: CellDefinition.(CellDefinition) -> CellDefinition): RowDefinition {
    return this.cell(column) {
        cellDefinition(it, it)
    }
}
fun RowDefinition.group(insideGroupDefinition: RowDefinition.(RowDefinition) -> RowDefinition): RowDefinition {
    return this.group {
        insideGroupDefinition(it, it)
    }
}
fun RowDefinition.collapse(insideGroupDefinition: RowDefinition.(RowDefinition) -> RowDefinition): RowDefinition {
    return this.collapse {
        insideGroupDefinition(it, it)
    }
}

/**
 * Creates new row in the spreadsheet.
 * @param rowDefinition closure defining the content of the row
 */
fun SheetDefinition.row(rowDefinition: RowDefinition.(RowDefinition) -> RowDefinition): SheetDefinition {
    return this.row {
        rowDefinition(it, it)
    }
}

/**
 * Creates new row in the spreadsheet.
 * @param row row number (1 based - the same as is shown in the file)
 * @param rowDefinition closure defining the content of the row
 */
fun SheetDefinition.row(row: Int, rowDefinition: RowDefinition.(RowDefinition) -> RowDefinition): SheetDefinition {
    return this.row(row) {
        rowDefinition(it, it)
    }
}

fun SheetDefinition.group(insideGroupDefinition: SheetDefinition.(SheetDefinition) -> SheetDefinition): SheetDefinition {
    return this.group {
        insideGroupDefinition(it, it)
    }
}
fun SheetDefinition.collapse(insideGroupDefinition: SheetDefinition.(SheetDefinition) -> SheetDefinition): SheetDefinition {
    return this.collapse {
        insideGroupDefinition(it, it)
    }
}

/**
 * Configures the basic page settings.
 * @param pageDefinition closure defining the page settings
 */
fun SheetDefinition.page(pageDefinition: PageDefinition.(PageDefinition) -> PageDefinition): SheetDefinition {
    return this.page {
        pageDefinition(it, it)
    }
}

fun SpreadsheetBuilder.build(workbookDefinition: WorkbookDefinition.(WorkbookDefinition) -> WorkbookDefinition) {
    return this.build {
        workbookDefinition(it, it)
    }
}

fun WorkbookDefinition.sheet(name: String, sheetDefinition: SheetDefinition.(SheetDefinition) -> SheetDefinition): WorkbookDefinition {
    return this.sheet(name) {
        sheetDefinition(it, it)
    }
}

fun CellCriterion.style(styleCriterion: CellStyleCriterion.(CellStyleCriterion) -> CellStyleCriterion): CellCriterion {
    return this.style {
        styleCriterion(it, it)
    }
}
fun CellCriterion.or(sheetCriterion: CellCriterion.(CellCriterion) -> CellCriterion): CellCriterion {
    return this.or {
        sheetCriterion(it, it)
    }
}


fun CellStyleCriterion.font(fontCriterion: FontCriterion.(FontCriterion) -> FontCriterion): CellStyleCriterion {
    return this.font {
        fontCriterion(it, it)
    }
}

/**
 * Configures all the borders of the cell.
 * @param borderConfiguration border configuration closure
 */
fun CellStyleCriterion.border(borderConfiguration: BorderCriterion.(BorderCriterion) -> BorderCriterion): CellStyleCriterion {
    return this.border {
        borderConfiguration(it, it)
    }
}

/**
 * Configures one border of the cell.
 * @param location border to be configured
 * @param borderConfiguration border configuration closure
 */
fun CellStyleCriterion.border(location: Keywords.BorderSide, borderConfiguration: BorderCriterion.(BorderCriterion) -> BorderCriterion): CellStyleCriterion {
    return this.border(location) {
        borderConfiguration(it, it)
    }
}

/**
 * Configures two borders of the cell.
 * @param first first border to be configured
 * @param second second border to be configured
 * @param borderConfiguration border configuration closure
 */
fun CellStyleCriterion.border(first: Keywords.BorderSide, second: Keywords.BorderSide, borderConfiguration: BorderCriterion.(BorderCriterion) -> BorderCriterion): CellStyleCriterion {
    return this.border(first, second) {
        borderConfiguration(it, it)
    }
}

/**
 * Configures three borders of the cell.
 * @param first first border to be configured
 * @param second second border to be configured
 * @param third third border to be configured
 * @param borderConfiguration border configuration closure
 */
fun CellStyleCriterion.border(first: Keywords.BorderSide, second: Keywords.BorderSide, third: Keywords.BorderSide, borderConfiguration: BorderCriterion.(BorderCriterion) -> BorderCriterion): CellStyleCriterion {
    return this.border(first, second, third) {
        borderConfiguration(it, it)
    }
}

fun RowCriterion.cell(cellCriterion: CellCriterion.(CellCriterion) -> CellCriterion): RowCriterion {
    return this.cell {
        cellCriterion(it, it)
    }
}
fun RowCriterion.cell(column: Int, cellCriterion: CellCriterion.(CellCriterion) -> CellCriterion): RowCriterion {
    return this.cell(column) {
        cellCriterion(it, it)
    }
}
fun RowCriterion.cell(column: String, cellCriterion: CellCriterion.(CellCriterion) -> CellCriterion): RowCriterion {
    return this.cell(column) {
        cellCriterion(it, it)
    }
}
fun RowCriterion.cell(from: Int, to: Int, cellCriterion: CellCriterion.(CellCriterion) -> CellCriterion): RowCriterion {
    return this.cell(from, to) {
        cellCriterion(it, it)
    }
}
fun RowCriterion.cell(from: String, to: String, cellCriterion: CellCriterion.(CellCriterion) -> CellCriterion): RowCriterion {
    return this.cell(from, to) {
        cellCriterion(it, it)
    }
}

fun RowCriterion.or(rowCriterion: RowCriterion.(RowCriterion) -> RowCriterion): RowCriterion {
    return this.or {
        rowCriterion(it, it)
    }
}

fun SheetCriterion.row(rowCriterion: RowCriterion.(RowCriterion) -> RowCriterion): SheetCriterion {
    return this.row {
        rowCriterion(it, it)
    }
}
fun SheetCriterion.row(row: Int, rowCriterion: RowCriterion.(RowCriterion) -> RowCriterion): SheetCriterion {
    return this.row(row) {
        rowCriterion(it, it)
    }
}
fun SheetCriterion.page(pageCriterion: PageCriterion.(PageCriterion) -> PageCriterion): SheetCriterion {
    return this.page {
        pageCriterion(it, it)
    }
}
fun SheetCriterion.or(sheetCriterion: SheetCriterion.(SheetCriterion) -> SheetCriterion): SheetCriterion {
    return this.or {
        sheetCriterion(it, it)
    }
}

fun SpreadsheetCriteria.query(workbookCriterion: WorkbookCriterion.(WorkbookCriterion) -> WorkbookCriterion): SpreadsheetCriteriaResult {
    return this.query {
        workbookCriterion(it, it)
    }
}
fun SpreadsheetCriteria.find(workbookCriterion: WorkbookCriterion.(WorkbookCriterion) -> WorkbookCriterion): Cell {
    return this.find {
        workbookCriterion(it, it)
    }
}
fun SpreadsheetCriteria.exists(workbookCriterion: WorkbookCriterion.(WorkbookCriterion) -> WorkbookCriterion): Boolean {
    return this.exists {
        workbookCriterion(it, it)
    }
}

fun WorkbookCriterion.sheet(name: String, sheetCriterion: SheetCriterion.(SheetCriterion) -> SheetCriterion): WorkbookCriterion {
    return this.sheet(name) {
        sheetCriterion(it, it)
    }
}
fun WorkbookCriterion.sheet(sheetCriterion: SheetCriterion.(SheetCriterion) -> SheetCriterion): WorkbookCriterion {
    return this.sheet {
        sheetCriterion(it, it)
    }
}
fun WorkbookCriterion.or(workbookCriterion: WorkbookCriterion.(WorkbookCriterion) -> WorkbookCriterion): WorkbookCriterion {
    return this.or {
        workbookCriterion(it, it)
    }
}

val DimensionModifier.cm: CellDefinition
    get() = this.cm()


val DimensionModifier.inch: CellDefinition
    get() = this.inch()


val DimensionModifier.inches: CellDefinition
    get() = this.inches()


val DimensionModifier.points: CellDefinition
    get() = this.points()


val BorderStyleProvider.none: BorderStyle
    get() = BorderStyle.NONE


val BorderStyleProvider.thin: BorderStyle
    get() = BorderStyle.THIN


val BorderStyleProvider.medium: BorderStyle
    get() = BorderStyle.MEDIUM


val BorderStyleProvider.dashed: BorderStyle
    get() = BorderStyle.DASHED


val BorderStyleProvider.dotted: BorderStyle
    get() = BorderStyle.DOTTED


val BorderStyleProvider.thick: BorderStyle
    get() = BorderStyle.THICK


val BorderStyleProvider.double: BorderStyle
    get() = BorderStyle.DOUBLE


val BorderStyleProvider.hair: BorderStyle
    get() = BorderStyle.HAIR


val BorderStyleProvider.mediumDashed: BorderStyle
    get() = BorderStyle.MEDIUM_DASHED


val BorderStyleProvider.dashDot: BorderStyle
    get() = BorderStyle.DASH_DOT


val BorderStyleProvider.mediumDashDot: BorderStyle
    get() = BorderStyle.MEDIUM_DASH_DOT


val BorderStyleProvider.dashDotDot: BorderStyle
    get() = BorderStyle.DASH_DOT_DOT


val BorderStyleProvider.mediumDashDotDot: BorderStyle
    get() = BorderStyle.MEDIUM_DASH_DOT_DOT


val BorderStyleProvider.slantedDashDot: BorderStyle
    get() = BorderStyle.SLANTED_DASH_DOT


val CellDefinition.auto: Keywords.Auto
    get() = Keywords.Auto.AUTO


val CellDefinition.to: Keywords.To
    get() = Keywords.To.TO


val CellDefinition.image: Keywords.Image
    get() = Keywords.Image.IMAGE


val ForegroundFillProvider.noFill: ForegroundFill
    get() = ForegroundFill.NO_FILL


val ForegroundFillProvider.solidForeground: ForegroundFill
    get() = ForegroundFill.SOLID_FOREGROUND


val ForegroundFillProvider.fineDots: ForegroundFill
    get() = ForegroundFill.FINE_DOTS


val ForegroundFillProvider.altBars: ForegroundFill
    get() = ForegroundFill.ALT_BARS


val ForegroundFillProvider.sparseDots: ForegroundFill
    get() = ForegroundFill.SPARSE_DOTS


val ForegroundFillProvider.thickHorizontalBands: ForegroundFill
    get() = ForegroundFill.THICK_HORZ_BANDS


val ForegroundFillProvider.thickVerticalBands: ForegroundFill
    get() = ForegroundFill.THICK_VERT_BANDS


val ForegroundFillProvider.thickBackwardDiagonals: ForegroundFill
    get() = ForegroundFill.THICK_BACKWARD_DIAG


val ForegroundFillProvider.thickForwardDiagonals: ForegroundFill
    get() = ForegroundFill.THICK_FORWARD_DIAG


val ForegroundFillProvider.bigSpots: ForegroundFill
    get() = ForegroundFill.BIG_SPOTS


val ForegroundFillProvider.bricks: ForegroundFill
    get() = ForegroundFill.BRICKS


val ForegroundFillProvider.thinHorizontalBands: ForegroundFill
    get() = ForegroundFill.THIN_HORZ_BANDS


val ForegroundFillProvider.thinVerticalBands: ForegroundFill
    get() = ForegroundFill.THIN_VERT_BANDS


val ForegroundFillProvider.thinBackwardDiagonals: ForegroundFill
    get() = ForegroundFill.THIN_BACKWARD_DIAG


val ForegroundFillProvider.thinForwardDiagonals: ForegroundFill
    get() = ForegroundFill.THICK_FORWARD_DIAG


val ForegroundFillProvider.squares: ForegroundFill
    get() = ForegroundFill.SQUARES


val ForegroundFillProvider.diamonds: ForegroundFill
    get() = ForegroundFill.DIAMONDS


val CellStyleDefinition.left: Keywords.BorderSideAndHorizontalAlignment
    get() = Keywords.BorderSideAndHorizontalAlignment.BSAHA_LEFT


val CellStyleDefinition.right: Keywords.BorderSideAndHorizontalAlignment
    get() = Keywords.BorderSideAndHorizontalAlignment.BSAHA_RIGHT


val CellStyleDefinition.top: Keywords.BorderSideAndVerticalAlignment
    get() = Keywords.BorderSideAndVerticalAlignment.BSAVA_TOP


val CellStyleDefinition.bottom: Keywords.BorderSideAndVerticalAlignment
    get() = Keywords.BorderSideAndVerticalAlignment.BSAVA_BOTTOM


val CellStyleDefinition.center: Keywords.VerticalAndHorizontalAlignment
    get() = Keywords.VerticalAndHorizontalAlignment.VAHA_CENTER


val CellStyleDefinition.justify: Keywords.VerticalAndHorizontalAlignment
    get() = Keywords.VerticalAndHorizontalAlignment.VAHA_JUSTIFY


val CellStyleDefinition.distributed: Keywords.PureVerticalAlignment
    get() = Keywords.PureVerticalAlignment.PVA_DISTRIBUTED


val CellStyleDefinition.Text: Keywords.Text
    get() = Keywords.Text.WRAP


val PageSettingsProvider.Portrait: Keywords.Orientation
    get() = Keywords.Orientation.PORTRAIT


val PageSettingsProvider.Landscape: Keywords.Orientation
    get() = Keywords.Orientation.LANDSCAPE


val PageSettingsProvider.Width: Keywords.Fit
    get() = Keywords.Fit.WIDTH


val PageSettingsProvider.Height: Keywords.Fit
    get() = Keywords.Fit.HEIGHT


val PageSettingsProvider.To: Keywords.To
    get() = Keywords.To.TO


val PageSettingsProvider.Letter: Keywords.Paper
    get() = Keywords.Paper.LETTER


val PageSettingsProvider.LetterSmall: Keywords.Paper
    get() = Keywords.Paper.LETTER_SMALL


val PageSettingsProvider.Tabloid: Keywords.Paper
    get() = Keywords.Paper.TABLOID


val PageSettingsProvider.Ledger: Keywords.Paper
    get() = Keywords.Paper.LEDGER


val PageSettingsProvider.Legal: Keywords.Paper
    get() = Keywords.Paper.LEGAL


val PageSettingsProvider.Statement: Keywords.Paper
    get() = Keywords.Paper.STATEMENT


val PageSettingsProvider.Executive: Keywords.Paper
    get() = Keywords.Paper.EXECUTIVE


val PageSettingsProvider.A3: Keywords.Paper
    get() = Keywords.Paper.A3


val PageSettingsProvider.A4: Keywords.Paper
    get() = Keywords.Paper.A4


val PageSettingsProvider.A4Small: Keywords.Paper
    get() = Keywords.Paper.A4_SMALL


val PageSettingsProvider.A5: Keywords.Paper
    get() = Keywords.Paper.A5


val PageSettingsProvider.B4: Keywords.Paper
    get() = Keywords.Paper.B4


val PageSettingsProvider.B5: Keywords.Paper
    get() = Keywords.Paper.B5


val PageSettingsProvider.Folio: Keywords.Paper
    get() = Keywords.Paper.FOLIO


val PageSettingsProvider.Quarto: Keywords.Paper
    get() = Keywords.Paper.QUARTO


val PageSettingsProvider.Standard10x14: Keywords.Paper
    get() = Keywords.Paper.STANDARD_10_14


val PageSettingsProvider.Standard11x17: Keywords.Paper
    get() = Keywords.Paper.STANDARD_11_17


val FontStylesProvider.Italic: FontStyle
    get() = FontStyle.ITALIC


val FontStylesProvider.Bold: FontStyle
    get() = FontStyle.BOLD


val FontStylesProvider.Strikeout: FontStyle
    get() = FontStyle.STRIKEOUT


val FontStylesProvider.Underline: FontStyle
    get() = FontStyle.UNDERLINE


val BorderPositionProvider.Left: Keywords.BorderSideAndHorizontalAlignment
    get() = Keywords.BorderSideAndHorizontalAlignment.BSAHA_LEFT;


val BorderPositionProvider.Right: Keywords.BorderSideAndHorizontalAlignment
    get() = Keywords.BorderSideAndHorizontalAlignment.BSAHA_RIGHT;


val BorderPositionProvider.Top: Keywords.BorderSideAndVerticalAlignment
    get() = Keywords.BorderSideAndVerticalAlignment.BSAVA_TOP;


val BorderPositionProvider.Bottom: Keywords.BorderSideAndVerticalAlignment
    get() = Keywords.BorderSideAndVerticalAlignment.BSAVA_BOTTOM;


val CellStyleDefinition.General: Keywords.PureHorizontalAlignment
    get() = Keywords.PureHorizontalAlignment.PHA_GENERAL

val CellStyleDefinition.Fill: Keywords.PureHorizontalAlignment
    get() = Keywords.PureHorizontalAlignment.PHA_FILL

val CellStyleDefinition.CenterSelection: Keywords.PureHorizontalAlignment
    get() = Keywords.PureHorizontalAlignment.PHA_CENTER_SELECTION

val SheetDefinition.Auto: Keywords.Auto
    get() = Keywords.Auto.AUTO


val ColorProvider.aliceBlue: Color
 get() = Color.aliceBlue

val ColorProvider.antiqueWhite: Color
 get() = Color.antiqueWhite

val ColorProvider.aqua: Color
 get() = Color.aqua

val ColorProvider.aquamarine: Color
 get() = Color.aquamarine

val ColorProvider.azure: Color
 get() = Color.azure

val ColorProvider.beige: Color
 get() = Color.beige

val ColorProvider.bisque: Color
 get() = Color.bisque

val ColorProvider.black: Color
 get() = Color.black

val ColorProvider.blanchedAlmond: Color
 get() = Color.blanchedAlmond

val ColorProvider.blue: Color
 get() = Color.blue

val ColorProvider.blueViolet: Color
 get() = Color.blueViolet

val ColorProvider.brown: Color
 get() = Color.brown

val ColorProvider.burlyWood: Color
 get() = Color.burlyWood

val ColorProvider.cadetBlue: Color
 get() = Color.cadetBlue

val ColorProvider.chartreuse: Color
 get() = Color.chartreuse

val ColorProvider.chocolate: Color
 get() = Color.chocolate

val ColorProvider.coral: Color
 get() = Color.coral

val ColorProvider.cornflowerBlue: Color
 get() = Color.cornflowerBlue

val ColorProvider.cornsilk: Color
 get() = Color.cornsilk

val ColorProvider.crimson: Color
 get() = Color.crimson

val ColorProvider.cyan: Color
 get() = Color.cyan

val ColorProvider.darkBlue: Color
 get() = Color.darkBlue

val ColorProvider.darkCyan: Color
 get() = Color.darkCyan

val ColorProvider.darkGoldenRod: Color
 get() = Color.darkGoldenRod

val ColorProvider.darkGray: Color
 get() = Color.darkGray

val ColorProvider.darkGreen: Color
 get() = Color.darkGreen

val ColorProvider.darkKhaki: Color
 get() = Color.darkKhaki

val ColorProvider.darkMagenta: Color
 get() = Color.darkMagenta

val ColorProvider.darkOliveGreen: Color
 get() = Color.darkOliveGreen

val ColorProvider.darkOrange: Color
 get() = Color.darkOrange

val ColorProvider.darkOrchid: Color
 get() = Color.darkOrchid

val ColorProvider.darkRed: Color
 get() = Color.darkRed

val ColorProvider.darkSalmon: Color
 get() = Color.darkSalmon

val ColorProvider.darkSeaGreen: Color
 get() = Color.darkSeaGreen

val ColorProvider.darkSlateBlue: Color
 get() = Color.darkSlateBlue

val ColorProvider.darkSlateGray: Color
 get() = Color.darkSlateGray

val ColorProvider.darkTurquoise: Color
 get() = Color.darkTurquoise

val ColorProvider.darkViolet: Color
 get() = Color.darkViolet

val ColorProvider.deepPink: Color
 get() = Color.deepPink

val ColorProvider.deepSkyBlue: Color
 get() = Color.deepSkyBlue

val ColorProvider.dimGray: Color
 get() = Color.dimGray

val ColorProvider.dodgerBlue: Color
 get() = Color.dodgerBlue

val ColorProvider.fireBrick: Color
 get() = Color.fireBrick

val ColorProvider.floralWhite: Color
 get() = Color.floralWhite

val ColorProvider.forestGreen: Color
 get() = Color.forestGreen

val ColorProvider.fuchsia: Color
 get() = Color.fuchsia

val ColorProvider.gainsboro: Color
 get() = Color.gainsboro

val ColorProvider.ghostWhite: Color
 get() = Color.ghostWhite

val ColorProvider.gold: Color
 get() = Color.gold

val ColorProvider.goldenRod: Color
 get() = Color.goldenRod

val ColorProvider.gray: Color
 get() = Color.gray

val ColorProvider.green: Color
 get() = Color.green

val ColorProvider.greenYellow: Color
 get() = Color.greenYellow

val ColorProvider.honeyDew: Color
 get() = Color.honeyDew

val ColorProvider.hotPink: Color
 get() = Color.hotPink

val ColorProvider.indianRed : Color
 get() = Color.indianRed

val ColorProvider.indigo : Color
 get() = Color.indigo

val ColorProvider.ivory: Color
 get() = Color.ivory

val ColorProvider.khaki: Color
 get() = Color.khaki

val ColorProvider.lavender: Color
 get() = Color.lavender

val ColorProvider.lavenderBlush: Color
 get() = Color.lavenderBlush

val ColorProvider.lawnGreen: Color
 get() = Color.lawnGreen

val ColorProvider.lemonChiffon: Color
 get() = Color.lemonChiffon

val ColorProvider.lightBlue: Color
 get() = Color.lightBlue

val ColorProvider.lightCoral: Color
 get() = Color.lightCoral

val ColorProvider.lightCyan: Color
 get() = Color.lightCyan

val ColorProvider.lightGoldenRodYellow: Color
 get() = Color.lightGoldenRodYellow

val ColorProvider.lightGray: Color
 get() = Color.lightGray

val ColorProvider.lightGreen: Color
 get() = Color.lightGreen

val ColorProvider.lightPink: Color
 get() = Color.lightPink

val ColorProvider.lightSalmon: Color
 get() = Color.lightSalmon

val ColorProvider.lightSeaGreen: Color
 get() = Color.lightSeaGreen

val ColorProvider.lightSkyBlue: Color
 get() = Color.lightSkyBlue

val ColorProvider.lightSlateGray: Color
 get() = Color.lightSlateGray

val ColorProvider.lightSteelBlue: Color
 get() = Color.lightSteelBlue

val ColorProvider.lightYellow: Color
 get() = Color.lightYellow

val ColorProvider.lime: Color
 get() = Color.lime

val ColorProvider.limeGreen: Color
 get() = Color.limeGreen

val ColorProvider.linen: Color
 get() = Color.linen

val ColorProvider.magenta: Color
 get() = Color.magenta

val ColorProvider.maroon: Color
 get() = Color.maroon

val ColorProvider.mediumAquaMarine: Color
 get() = Color.mediumAquaMarine

val ColorProvider.mediumBlue: Color
 get() = Color.mediumBlue

val ColorProvider.mediumOrchid: Color
 get() = Color.mediumOrchid

val ColorProvider.mediumPurple: Color
 get() = Color.mediumPurple

val ColorProvider.mediumSeaGreen: Color
 get() = Color.mediumSeaGreen

val ColorProvider.mediumSlateBlue: Color
 get() = Color.mediumSlateBlue

val ColorProvider.mediumSpringGreen: Color
 get() = Color.mediumSpringGreen

val ColorProvider.mediumTurquoise: Color
 get() = Color.mediumTurquoise

val ColorProvider.mediumVioletRed: Color
 get() = Color.mediumVioletRed

val ColorProvider.midnightBlue: Color
 get() = Color.midnightBlue

val ColorProvider.mintCream: Color
 get() = Color.mintCream

val ColorProvider.mistyRose: Color
 get() = Color.mistyRose

val ColorProvider.moccasin: Color
 get() = Color.moccasin

val ColorProvider.navajoWhite: Color
 get() = Color.navajoWhite

val ColorProvider.navy: Color
 get() = Color.navy

val ColorProvider.oldLace: Color
 get() = Color.oldLace

val ColorProvider.olive: Color
 get() = Color.olive

val ColorProvider.oliveDrab: Color
 get() = Color.oliveDrab

val ColorProvider.orange: Color
 get() = Color.orange

val ColorProvider.orangeRed: Color
 get() = Color.orangeRed

val ColorProvider.orchid: Color
 get() = Color.orchid

val ColorProvider.paleGoldenRod: Color
 get() = Color.paleGoldenRod

val ColorProvider.paleGreen: Color
 get() = Color.paleGreen

val ColorProvider.paleTurquoise: Color
 get() = Color.paleTurquoise

val ColorProvider.paleVioletRed: Color
 get() = Color.paleVioletRed

val ColorProvider.papayaWhip: Color
 get() = Color.papayaWhip

val ColorProvider.peachPuff: Color
 get() = Color.peachPuff

val ColorProvider.peru: Color
 get() = Color.peru

val ColorProvider.pink: Color
 get() = Color.pink

val ColorProvider.plum: Color
 get() = Color.plum

val ColorProvider.powderBlue: Color
 get() = Color.powderBlue

val ColorProvider.purple: Color
 get() = Color.purple

val ColorProvider.rebeccaPurple: Color
 get() = Color.rebeccaPurple

val ColorProvider.red: Color
 get() = Color.red

val ColorProvider.rosyBrown: Color
 get() = Color.rosyBrown

val ColorProvider.royalBlue: Color
 get() = Color.royalBlue

val ColorProvider.saddleBrown: Color
 get() = Color.saddleBrown

val ColorProvider.salmon: Color
 get() = Color.salmon

val ColorProvider.sandyBrown: Color
 get() = Color.sandyBrown

val ColorProvider.seaGreen: Color
 get() = Color.seaGreen

val ColorProvider.seaShell: Color
 get() = Color.seaShell

val ColorProvider.sienna: Color
 get() = Color.sienna

val ColorProvider.silver: Color
 get() = Color.silver

val ColorProvider.skyBlue: Color
 get() = Color.skyBlue

val ColorProvider.slateBlue: Color
 get() = Color.slateBlue

val ColorProvider.slateGray: Color
 get() = Color.slateGray

val ColorProvider.snow: Color
 get() = Color.snow

val ColorProvider.springGreen: Color
 get() = Color.springGreen

val ColorProvider.steelBlue: Color
 get() = Color.steelBlue

val ColorProvider.tan: Color
 get() = Color.tan

val ColorProvider.teal: Color
 get() = Color.teal

val ColorProvider.thistle: Color
 get() = Color.thistle

val ColorProvider.tomato: Color
 get() = Color.tomato

val ColorProvider.turquoise: Color
 get() = Color.turquoise

val ColorProvider.violet: Color
 get() = Color.violet

val ColorProvider.wheat: Color
 get() = Color.wheat

val ColorProvider.white: Color
 get() = Color.white

val ColorProvider.whiteSmoke: Color
 get() = Color.whiteSmoke

val ColorProvider.yellow: Color
 get() = Color.yellow

val ColorProvider.yellowGreen: Color
 get() = Color.yellowGreen


val SheetStateProvider.locked: Keywords.SheetState
 get() = Keywords.SheetState.LOCKED

val SheetStateProvider.visible: Keywords.SheetState
 get() = Keywords.SheetState.VISIBLE

val SheetStateProvider.hidden: Keywords.SheetState
 get() = Keywords.SheetState.HIDDEN

val SheetStateProvider.veryHidden: Keywords.SheetState
 get() = Keywords.SheetState.VERY_HIDDEN

