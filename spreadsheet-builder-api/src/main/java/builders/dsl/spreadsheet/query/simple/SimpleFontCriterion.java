package builders.dsl.spreadsheet.query.simple;

import builders.dsl.spreadsheet.api.*;
import builders.dsl.spreadsheet.query.api.FontCriterion;
import builders.dsl.spreadsheet.query.api.Predicate;

import java.util.EnumSet;

final class SimpleFontCriterion implements FontCriterion {

    private final SimpleCellCriterion parent;

    SimpleFontCriterion(SimpleCellCriterion parent) {
        this.parent = parent;
    }

    @Override
    public SimpleFontCriterion color(String hexColor) {
        color(new Color(hexColor));
        return this;
    }

    @Override
    public SimpleFontCriterion color(final Color color) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && color.equals(font.getColor());
            }
        });
        return this;
    }

    @Override
    public SimpleFontCriterion color(final Predicate<Color> conition) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && conition.test(font.getColor());
            }
        });
        return this;
    }

    @Override
    public SimpleFontCriterion size(final int size) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && size == font.getSize();
            }
        });
        return this;
    }

    @Override
    public SimpleFontCriterion size(final Predicate<Integer> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && predicate.test(font.getSize());
            }
        });
        return this;
    }

    @Override
    public SimpleFontCriterion name(final String name) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && name.equals(font.getName());
            }
        });
        return this;
    }

    @Override
    public SimpleFontCriterion name(final Predicate<String> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && predicate.test(font.getName());
            }
        });
        return this;
    }

    @Override
    public SimpleFontCriterion style(final FontStyle first, final FontStyle... other) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                if (font == null) {
                    return false;
                }

                EnumSet<FontStyle> wanted = EnumSet.of(first, other);
                EnumSet<FontStyle> actual = font.getStyles();

                for (FontStyle fs : wanted) {
                    if (!actual.contains(fs)) {
                        return false;
                    }
                }

                return true;
            }
        });
        return this;
    }

    @Override
    public SimpleFontCriterion style(final Predicate<EnumSet<FontStyle>> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && predicate.test(font.getStyles());
            }
        });
        return this;
    }

    @Override
    public FontCriterion having(final Predicate<Font> fontPredicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && fontPredicate.test(font);
            }
        });
        return this;
    }
}
