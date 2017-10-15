package builders.dsl.spreadsheet.api;

public final class Keywords {
    //CHECKSTYLE:OFF
    public static final BorderStyle none = BorderStyle.NONE;
    public static final BorderStyle thin = BorderStyle.THIN;
    public static final BorderStyle medium = BorderStyle.MEDIUM;
    public static final BorderStyle dashed = BorderStyle.DASHED;
    public static final BorderStyle dotted = BorderStyle.DOTTED;
    public static final BorderStyle thick = BorderStyle.THICK;
    public static final BorderStyle doubleBorder = BorderStyle.DOUBLE;
    public static final BorderStyle hair = BorderStyle.HAIR;
    public static final BorderStyle mediumDashed = BorderStyle.MEDIUM_DASHED;
    public static final BorderStyle mashDot = BorderStyle.DASH_DOT;
    public static final BorderStyle mediumDashDot = BorderStyle.MEDIUM_DASH_DOT;
    public static final BorderStyle dashDotDot = BorderStyle.DASH_DOT_DOT;
    public static final BorderStyle mediumDashDotDot = BorderStyle.MEDIUM_DASH_DOT_DOT;
    public static final BorderStyle slantedDashDot = BorderStyle.SLANTED_DASH_DOT;
    public static final Keywords.Orientation portrait = Keywords.Orientation.PORTRAIT;
    public static final Keywords.Orientation landscape = Keywords.Orientation.LANDSCAPE;
    public static final Keywords.Fit width = Keywords.Fit.WIDTH;
    public static final Keywords.Fit height = Keywords.Fit.HEIGHT;
    public static final Keywords.Paper letter = Keywords.Paper.LETTER;
    public static final Keywords.Paper letterSmall = Keywords.Paper.LETTER_SMALL;
    public static final Keywords.Paper tabloid = Keywords.Paper.TABLOID;
    public static final Keywords.Paper ledger = Keywords.Paper.LEDGER;
    public static final Keywords.Paper legal = Keywords.Paper.LEGAL;
    public static final Keywords.Paper statement = Keywords.Paper.STATEMENT;
    public static final Keywords.Paper executive = Keywords.Paper.EXECUTIVE;
    public static final Keywords.Paper A3 = Keywords.Paper.A3;
    public static final Keywords.Paper A4 = Keywords.Paper.A4;
    public static final Keywords.Paper A4Small = Keywords.Paper.A4_SMALL;
    public static final Keywords.Paper A5 = Keywords.Paper.A5;
    public static final Keywords.Paper B4 = Keywords.Paper.B4;
    public static final Keywords.Paper B5 = Keywords.Paper.B5;
    public static final Keywords.Paper folio = Keywords.Paper.FOLIO;
    public static final Keywords.Paper quarto = Keywords.Paper.QUARTO;
    public static final Keywords.Paper standard10x14 = Keywords.Paper.STANDARD_10_14;
    public static final Keywords.Paper standard11x17 = Keywords.Paper.STANDARD_11_17;
    public static final FontStyle italic = FontStyle.ITALIC;
    public static final FontStyle bold = FontStyle.BOLD;
    public static final FontStyle strikeout = FontStyle.STRIKEOUT;
    public static final FontStyle underline = FontStyle.UNDERLINE;
    public static final ForegroundFill noFill = ForegroundFill.NO_FILL;
    public static final ForegroundFill solidForeground = ForegroundFill.SOLID_FOREGROUND;
    public static final ForegroundFill fineDots = ForegroundFill.FINE_DOTS;
    public static final ForegroundFill altBars = ForegroundFill.ALT_BARS;
    public static final ForegroundFill sparseDots = ForegroundFill.SPARSE_DOTS;
    public static final ForegroundFill thickHorizontalBands = ForegroundFill.THICK_HORZ_BANDS;
    public static final ForegroundFill thickVerticalBands = ForegroundFill.THICK_VERT_BANDS;
    public static final ForegroundFill thickBackwardDiagonals = ForegroundFill.THICK_BACKWARD_DIAG;
    public static final ForegroundFill thickForwardDiagonals = ForegroundFill.THICK_FORWARD_DIAG;
    public static final ForegroundFill bigSpots = ForegroundFill.BIG_SPOTS;
    public static final ForegroundFill bricks = ForegroundFill.BRICKS;
    public static final ForegroundFill thinHorizontalBands = ForegroundFill.THIN_HORZ_BANDS;
    public static final ForegroundFill thinVerticalBands = ForegroundFill.THIN_VERT_BANDS;
    public static final ForegroundFill thinBackwardDiagonals = ForegroundFill.THIN_BACKWARD_DIAG;
    public static final ForegroundFill thinForwardDiagonals = ForegroundFill.THICK_FORWARD_DIAG;
    public static final ForegroundFill squares = ForegroundFill.SQUARES;
    public static final ForegroundFill diamonds = ForegroundFill.DIAMONDS;
    public static final BorderSideAndHorizontalAlignment left = BorderSideAndHorizontalAlignment.BSAHA_LEFT;
    public static final BorderSideAndHorizontalAlignment right = BorderSideAndHorizontalAlignment.BSAHA_RIGHT;
    public static final Keywords.BorderSideAndVerticalAlignment top = BorderSideAndVerticalAlignment.BSAVA_TOP;
    public static final Keywords.BorderSideAndVerticalAlignment bottom = BorderSideAndVerticalAlignment.BSAVA_BOTTOM;
    public static final Keywords.VerticalAndHorizontalAlignment center = VerticalAndHorizontalAlignment.VAHA_CENTER;
    public static final Keywords.VerticalAndHorizontalAlignment justify = VerticalAndHorizontalAlignment.VAHA_JUSTIFY;
    public static final Keywords.PureVerticalAlignment distributed = Keywords.PureVerticalAlignment.PVA_DISTRIBUTED;
    public static final Keywords.Text text = Keywords.Text.WRAP;
    public static final Keywords.Auto auto = Keywords.Auto.AUTO;
    public static final Keywords.To to = Keywords.To.TO;
    public static final Keywords.Image image = Keywords.Image.IMAGE;
    public static final PureHorizontalAlignment general = PureHorizontalAlignment.PHA_GENERAL;
    public static final PureHorizontalAlignment fill = PureHorizontalAlignment.PHA_FILL;
    public static final PureHorizontalAlignment centerSelection = PureHorizontalAlignment.PHA_CENTER_SELECTION;
    public static final Keywords.SheetState locked = SheetState.LOCKED;
    public static final Keywords.SheetState visible = SheetState.VISIBLE;
    public static final Keywords.SheetState hidden = SheetState.HIDDEN;
    public static final Keywords.SheetState veryHidden = SheetState.VERY_HIDDEN;


    private Keywords() {}

    public enum Text {
        WRAP
    }

    public enum To {
        TO
    }

    public enum Image {
        IMAGE
    }

    public enum Auto {
        AUTO
    }

    public enum BorderSideAndVerticalAlignment implements BorderSide, VerticalAlignment {
        BSAVA_TOP,
        BSAVA_BOTTOM
    }

    public enum BorderSideAndHorizontalAlignment implements BorderSide, HorizontalAlignment {
        BSAHA_LEFT,
        BSAHA_RIGHT
    }

    public enum VerticalAndHorizontalAlignment implements VerticalAlignment, HorizontalAlignment {
        VAHA_CENTER,
        VAHA_JUSTIFY
    }

    public enum PureVerticalAlignment implements VerticalAlignment {
        PVA_DISTRIBUTED
    }

    public enum Orientation {
        LANDSCAPE,
        PORTRAIT
    }

    public enum Fit {
        HEIGHT,
        WIDTH
    }

    public enum Paper {
        LETTER,
        LETTER_SMALL,
        TABLOID,
        LEDGER,
        LEGAL,
        STATEMENT,
        EXECUTIVE,
        A3,
        A4,
        A4_SMALL,
        A5,
        B4,
        B5,
        FOLIO,
        QUARTO,
        STANDARD_10_14,
        STANDARD_11_17
    }

    public enum PureHorizontalAlignment implements HorizontalAlignment {
        PHA_GENERAL,
        PHA_FILL,
        PHA_CENTER_SELECTION
    }

    public enum SheetState {
        LOCKED,
        VISIBLE,
        HIDDEN,
        VERY_HIDDEN
    }


    public interface BorderSide {
        BorderSide LEFT = BorderSideAndHorizontalAlignment.BSAHA_LEFT;
        BorderSide RIGHT = BorderSideAndHorizontalAlignment.BSAHA_RIGHT;
        BorderSide TOP = BorderSideAndVerticalAlignment.BSAVA_TOP;
        BorderSide BOTTOM = BorderSideAndVerticalAlignment.BSAVA_BOTTOM;

        BorderSide[] BORDER_SIDES = {TOP, BOTTOM, LEFT, RIGHT};
    }

    public interface VerticalAlignment {

        VerticalAlignment TOP = BorderSideAndVerticalAlignment.BSAVA_TOP;
        VerticalAlignment CENTER = VerticalAndHorizontalAlignment.VAHA_CENTER;
        VerticalAlignment BOTTOM = BorderSideAndVerticalAlignment.BSAVA_BOTTOM;
        VerticalAlignment JUSTIFY = VerticalAndHorizontalAlignment.VAHA_JUSTIFY;
        VerticalAlignment DISTRIBUTED = PureVerticalAlignment.PVA_DISTRIBUTED;

        VerticalAlignment[] VERTICAL_ALIGNMENTS = {TOP, CENTER, BOTTOM, JUSTIFY, DISTRIBUTED};

    }

    public interface HorizontalAlignment {
        HorizontalAlignment RIGHT = BorderSideAndHorizontalAlignment.BSAHA_RIGHT;
        HorizontalAlignment LEFT = BorderSideAndHorizontalAlignment.BSAHA_LEFT;
        HorizontalAlignment GENERAL = PureHorizontalAlignment.PHA_GENERAL;
        HorizontalAlignment CENTER = VerticalAndHorizontalAlignment.VAHA_CENTER;
        HorizontalAlignment FILL = PureHorizontalAlignment.PHA_FILL;
        HorizontalAlignment JUSTIFY = VerticalAndHorizontalAlignment.VAHA_JUSTIFY;
        HorizontalAlignment CENTER_SELECTION = PureHorizontalAlignment.PHA_CENTER_SELECTION;

        HorizontalAlignment[] HORIZONTAL_ALIGNMENTS = {RIGHT, LEFT, GENERAL, CENTER, FILL, JUSTIFY, CENTER_SELECTION};
    }
    //CHECKSTYLE:ON
}
