package builders.dsl.spreadsheet.api.kotlin.test

import builders.dsl.spreadsheet.api.Color.*
import builders.dsl.spreadsheet.api.DataRow
import builders.dsl.spreadsheet.api.ForegroundFill
import builders.dsl.spreadsheet.api.Keywords
import builders.dsl.spreadsheet.builder.api.CellDefinition
import builders.dsl.spreadsheet.builder.api.SpreadsheetBuilder
import builders.dsl.spreadsheet.builder.poi.PoiSpreadsheetBuilder
import builders.dsl.spreadsheet.query.poi.PoiSpreadsheetCriteria
import org.apache.poi.openxml4j.exceptions.InvalidFormatException
import org.junit.Assert.*
import org.junit.Rule
import org.junit.Test
import org.junit.rules.TemporaryFolder
import java.io.IOException
import java.util.*

import builders.dsl.spreadsheet.api.kotlin.*

class Java8BuilderTest {

    @Rule
    var tmp = TemporaryFolder()

    @Test
    @Throws(IOException::class)
    fun testBuilder() {
        val excel = tmp.newFile()
        PoiSpreadsheetBuilder.create(excel).build {
            sheet("test") {
                row {
                    cell { value("Hello World") }
                }
            }
        }

        assertEquals(1, PoiSpreadsheetCriteria.FACTORY.forFile(excel).query {
            sheet("test") {
                row {
                    cell { value("Hello World") }
                }
            }
        }.cells.size())

    }

    @Test
    @Throws(IOException::class, InvalidFormatException::class)
    fun testBuilderFull() {
        val today = Date()
        val excel = tmp.newFile()

        buildSpreadsheet(PoiSpreadsheetBuilder.create(excel), today)

        val matcher = PoiSpreadsheetCriteria.FACTORY.forFile(excel)

        val allCells = matcher.all()

        assertEquals(80130, allCells.getCells().size.toLong())
        assertEquals(21, allCells.getSheets().size.toLong())
        assertEquals(20065, allCells.getRows().size.toLong())


        val sampleCells = matcher.query({ w -> w.sheet("Sample") })

        assertEquals(2, sampleCells.getCells().size.toLong())
        assertEquals(1, sampleCells.getSheets().size.toLong())
        assertEquals(1, sampleCells.getRows().size.toLong())

        val rowCells = matcher.query({ w ->
            w.sheet("many rows", { s -> s.row(1) }
            )
        }
        )

        assertEquals(4, rowCells.getCells().size.toLong())
        assertEquals(1, rowCells.getSheets().size.toLong())
        assertEquals(1, rowCells.getRows().size.toLong())

        val manyRowsHeader = matcher.query({ w ->
            w.sheet("many rows", { s -> s.row(1) }
            )
        }
        ).getRow()

        assertNotNull(manyRowsHeader)

        val manyRowsDataRow = matcher.query({ w -> w.sheet("many rows", { s -> s.row(2) }) }).getRow()

        val dataRow = DataRow.create(manyRowsDataRow, manyRowsHeader)

        assertNotNull(dataRow.get("One"))
        assertEquals("1", dataRow.get("One").getValue())


        val dataRowMapping = HashMap<String, Int>()
        dataRowMapping.put("primo", 1)


        val dataRowFromMapping = DataRow.create(dataRowMapping, manyRowsDataRow)
        assertNotNull(dataRowFromMapping.get("primo"))
        assertEquals("1", dataRowFromMapping.get("primo").value)


        val someCells = matcher.query({ w ->
            w.sheet({ s ->
                s.row({ r ->
                    r.cell({ c -> c.date(today) }
                    )
                }
                )
            }
            )
        }
        )

        assertEquals(1, someCells.getCells().size.toLong())


        val commentedCells = matcher.query({ w ->
            w.sheet({ s ->
                s.row({ r ->
                    r.cell({ c -> c.comment("This is a date!") }
                    )
                }
                )
            }
            )
        }
        )

        assertEquals(1, commentedCells.getCells().size.toLong())


        val namedCells = matcher.query({ w -> w.sheet({ s -> s.row({ r -> r.cell({ c -> c.name("_Cell10") }) }) }) })

        assertEquals(1, namedCells.getCells().size.toLong())


        val dateCells = matcher.query({ w -> w.sheet({ s -> s.row({ r -> r.cell({ c -> c.style({ st -> st.format("d.m.y") }) }) }) }) })

        assertEquals(1, dateCells.getCells().size.toLong())

        val filledCells = matcher.query({ w -> w.sheet({ s -> s.row({ r -> r.cell({ c -> c.style({ st -> st.fill(Keywords.fineDots) }) }) }) }) })

        assertEquals(1, filledCells.getCells().size.toLong())

        val magentaCells = matcher.query({ w -> w.sheet({ s -> s.row({ r -> r.cell({ c -> c.style({ st -> st.foreground(aquamarine) }) }) }) }) })

        assertEquals(1, magentaCells.getCells().size.toLong())

        val redOnes = matcher.query({ w -> w.sheet({ s -> s.row({ r -> r.cell({ c -> c.style({ st -> st.font({ f -> f.color(red) }) }) }) }) }) })

        assertEquals(20006, redOnes.getCells().size.toLong())
        assertEquals(20004, redOnes.getRows().size.toLong())


        val boldOnes = matcher.query({ w -> w.sheet({ s -> s.row({ r -> r.cell({ c -> c.style({ st -> st.font({ f -> f.style(Keywords.bold) }) }) }) }) }) })

        assertEquals(5, boldOnes.getCells().size.toLong())

        val bigOnes = matcher.query({ w -> w.sheet({ s -> s.row({ r -> r.cell({ c -> c.style({ st -> st.font({ f -> f.size(22) }) }) }) }) }) })

        assertEquals(40002, bigOnes.getCells().size.toLong())

        val bordered = matcher.query({ w -> w.sheet({ s -> s.row({ r -> r.cell({ c -> c.style({ st -> st.border(Keywords.top, { b -> b.style(Keywords.thin) }) }) }) }) }) })

        assertEquals(10, bordered.getCells().size.toLong())

        val combined = matcher.query({ w ->
            w.sheet({ s ->
                s.row({ r ->
                    r.cell({ c ->
                        c.value("Bold Red 22")
                        c.style({ st -> st.font({ f -> f.color(red) }) })
                    })
                })
            })
        })

        assertEquals(1, combined.getCells().size.toLong())

        val conjunction = matcher.query({ w ->
            w.sheet({ s ->
                s.row({ r ->
                    r.or({ o ->
                        o.cell({ c -> c.value("Bold Red 22") })
                        o.cell({ c -> c.value("A") })
                    })
                })
            })
        })

        assertEquals(3, conjunction.getCells().size.toLong())

        val traversal = matcher.query({ w -> w.sheet("Traversal", { s -> s.row({ r -> r.cell({ c -> c.value("E") }) }) }) })

        assertEquals(1, traversal.getCells().size.toLong())

        val cellE = traversal.iterator().next()

        assertEquals("Traversal", cellE.getRow().getSheet().getName())
        assertNotNull(cellE.getRow().getSheet().getPrevious())
        assertEquals("Formula", cellE.getRow().getSheet().getPrevious().getName())
        assertNotNull(cellE.getRow().getSheet().getNext())
        assertEquals("Border", cellE.getRow().getSheet().getNext().getName())
        assertEquals(2, cellE.getRow().getNumber().toLong())
        assertNotNull(cellE.getRow().getAbove())
        assertEquals(1, cellE.getRow().getAbove().getNumber().toLong())
        assertNotNull(cellE.getRow().getBellow())
        assertEquals(3, cellE.getRow().getBellow().getNumber().toLong())
        assertEquals(2, cellE.getColspan().toLong())
        assertNotNull(cellE.getAboveLeft())
        assertEquals("A", cellE.getAboveLeft().getValue())
        assertNotNull(cellE.getAbove())
        assertEquals("B", cellE.getAbove().getValue())
        assertNotNull(cellE.getAboveRight())
        assertEquals("C", cellE.getAboveRight().getValue())
        assertNotNull(cellE.getLeft())
        assertEquals("D", cellE.getLeft().getValue())
        assertNotNull(cellE.getRight())
        assertEquals("F", cellE.getRight().getValue())
        assertNotNull(cellE.getBellowLeft())
        assertEquals("G", cellE.getBellowLeft().getValue())
        assertNotNull(cellE.getBellowRight())
        assertEquals("I", cellE.getBellowRight().getValue())
        assertNotNull(cellE.getBellow())
        assertEquals("H", cellE.getBellow().getValue())
        assertEquals("J", cellE.getBellow().getBellow().getValue())

        val zeroCells = matcher.query({ w -> w.sheet("Zero", { s -> s.row({ r -> r.cell({ c -> c.value(0) }) }) }) })

        assertEquals(1, zeroCells.getCells().size.toLong())
        assertEquals(0.0, zeroCells.getCell().getValue())

        val noneCell = matcher.find({ w -> w.sheet("Styles", { s -> s.row({ r -> r.cell({ c -> c.value("NONE") }) }) }) })

        val redCell = matcher.find({ w -> w.sheet("Styles", { s -> s.row({ r -> r.cell({ c -> c.value("RED") }) }) }) })

        val blueCell = matcher.find({ w -> w.sheet("Styles", { s -> s.row({ r -> r.cell({ c -> c.value("BLUE") }) }) }) })

        val greenCell = matcher.find({ w -> w.sheet("Styles", { s -> s.row({ r -> r.cell({ c -> c.value("GREEN") }) }) }) })

        assertNull(noneCell.getStyle().getForeground())
        assertEquals(red, redCell.getStyle().getForeground())
        assertEquals(blue, blueCell.getStyle().getForeground())
        assertEquals(green, greenCell.getStyle().getForeground())

        assertEquals(1, matcher.query({ w ->
            w.sheet({ s ->
                s.page({ p ->
                    p.paper(Keywords.Paper.A5)
                    p.orientation(Keywords.Orientation.LANDSCAPE)
                })
            })
        }).getCells().size.toLong())
        assertEquals(19, matcher.query({ w -> w.sheet({ s -> s.state(Keywords.visible) }) }).getSheets().size.toLong())
        assertEquals(1, matcher.query({ w -> w.sheet("Hidden", { s -> s.state(Keywords.hidden) }) }).getSheets().size.toLong())
        assertEquals(1, matcher.query({ w -> w.sheet("Very hidden", { s -> s.state(Keywords.veryHidden) }) }).getSheets().size.toLong())

        val tmpFile = tmp.newFile()

        PoiSpreadsheetBuilder.create(tmpFile, excel).build({ w ->
            w.sheet("Sample", { s ->
                s.row({ r ->
                    r.cell({ c -> c.value("Ahoj") })
                    r.cell({ c -> c.value("Svete") })
                })
            })
        })

        assertEquals(0, PoiSpreadsheetCriteria.FACTORY.forFile(tmpFile).query({ w ->
            w.sheet("Sample", { s ->
                s.row({ r ->
                    r.cell({ c -> c.value("Hello") })

                })
            })
        }).getCells().size())


        assertEquals(1, PoiSpreadsheetCriteria.FACTORY.forFile(tmpFile).query({ w ->
            w.sheet("Sample", { s ->
                s.row({ r ->
                    r.cell({ c -> c.value("Ahoj") })

                })
            })
        }).getCells().size())
    }

    private fun buildSpreadsheet(builder: SpreadsheetBuilder, today: Date) {
        builder.build { w ->
            w.style("red") { s -> s.font { f -> f.color(red) } }
            w.style("bold") { s -> s.font { f -> f.style(Keywords.bold) } }

            w.style("borders") { s ->
                s.font { f -> f.color(red) }
                s.border { b -> b.style(Keywords.thin).color(black) }
            }

            w.style("centered") { s -> s.align(Keywords.center, Keywords.center) }

            w.style("redfg") { s -> s.foreground(red) }

            w.style("greenfg") { s -> s.foreground(green) }

            w.style("bluefg") { s -> s.foreground(blue) }

            w.style("nonefg") { s -> }


            w.apply(MyStyles::class.java)

            w.sheet("many rows") { s ->
                s.filter(Keywords.auto)
                s.row { r ->
                    r.cell("One")
                    r.cell("Two")
                    r.cell("Three")
                    r.cell("Four")
                }
                for (i in 0..19999) {
                    s.row { r ->
                        r.cell { c ->
                            c.value("1")
                            c.styles("h1", "red")
                        }
                        r.cell { c ->
                            c.value("2")
                            c.style("h2")
                        }
                        r.cell { c ->
                            c.value("3")
                            c.style("h1")
                        }
                        r.cell { c ->
                            c.value("4")
                            c.style("h2")
                        }
                    }
                }
            }

            w.sheet("Sample") { s ->
                s.row { r ->
                    r.cell { c ->
                        c.value("Hello")
                        c.style("h1")
                    }
                    r.cell { c ->
                        c.value("World")
                        c.style("h2")
                    }
                }
            }

            w.sheet("Fonts") { s ->
                s.row { r ->
                    r.cell { c ->
                        c.width(Keywords.auto)
                        c.value("Bold Red 22")
                        c.style { st ->
                            st.font { f ->
                                f.style(Keywords.bold)
                                f.color(red)
                                f.size(22)
                            }
                        }
                    }
                    r.cell { c ->
                        c.width(Keywords.auto)
                        c.value("Underline Courier New")
                        c.style { st ->
                            st.font { f ->
                                f.style(Keywords.underline)
                                f.name("Courier New")
                            }
                        }
                    }
                    r.cell { c ->
                        c.width(Keywords.auto)
                        c.value("Italic")
                        c.style { st -> st.font { f -> f.style(Keywords.italic) } }
                    }
                    r.cell { c ->
                        c.width(Keywords.auto)
                        c.value("Strikeout")
                        c.style { st -> st.font { f -> f.style(Keywords.strikeout) } }
                    }
                }
            }

            w.sheet("Image") { s ->
                s.row(3) { r ->
                    r.cell("C") { c ->
                        //                         c.png(c.getImage()).from("https://goo.gl/UcL1wy");
                    }
                }
            }

            w.sheet("Rich Text") { s ->
                s.row { r ->
                    r.cell { c ->
                        c.text("Little")
                        c.text(" ")
                        c.text("Red") { f ->
                            f.color(red)
                            f.size(22)
                            f.name("Times New Roman")
                        }
                        c.text(" ")
                        c.text("Riding") { f ->
                            f.style(Keywords.italic)
                            f.size(18)
                        }
                        c.text(" ")
                        c.text("Hood") { f ->
                            f.style(Keywords.bold)
                            f.size(22)
                        }

                    }
                }
                s.row { r ->
                    r.cell { c ->
                        c.style { st -> st.wrap(Keywords.text) }
                        c.text("First Line")
                        c.text("\n")
                        c.text("Second Line") { f ->
                            f.style(Keywords.bold)
                            f.size(12)
                        }
                        c.text("\n")

                        val values = LinkedHashMap<String, String>()
                        values.put("foo", "bar")
                        values.put("boo", "cow")
                        values.put("empty", "")
                        values.put("", "nothing")

                        for (entry in values.entries) {
                            printMap(c, entry)
                        }
                        c.text("\n\n")
                        c.text("Next line after two spaces")
                        c.text("\n")
                        c.text("Last line") { f -> f.style(Keywords.italic) }
                        c.text("\n")
                    }
                }

                s.row { r ->
                    r.cell { c ->
                        c.text("Genomics England Consent Withdrawal Options")
                        c.text("\n\n")
                        c.text("Enumerations") { f ->
                            f.size(12)
                            f.style(Keywords.bold)
                        }
                        c.text("\n")

                        val withdrawals = LinkedHashMap<String, String>()
                        withdrawals.put("FULL_WITHDRAWAL", "OPTION 2: FULL WITHDRAWAL: No further use")
                        withdrawals.put("PARTIAL_WITHDRAWAL", "OPTION 1: PARTIAL WITHDRAWAL: No further contact")

                        for (entry in withdrawals.entries) {
                            printMap(c, entry)
                        }
                    }
                }
            }

            w.sheet("Cell Addressing") { s ->
                s.filter(Keywords.auto)
                s.row(2) { r ->
                    r.style { st ->
                        st.background(whiteSmoke)
                        st.border(Keywords.top, Keywords.bottom) { b ->
                            b.style(Keywords.thin)
                            b.color(black)
                        }
                    }
                    r.cell("B") { c ->
                        c.value("A")
                        c.style { st ->
                            st.border(Keywords.left) { b ->
                                b.style(Keywords.thin)
                                b.color(black)
                            }
                        }
                    }
                    r.cell("B")
                    r.cell { c ->
                        c.value("C")
                        c.style { st ->
                            st.border(Keywords.right) { b ->
                                b.style(Keywords.thin)
                                b.color(black)
                            }
                        }
                    }
                }
                s.row { r ->
                    r.cell("B") { c -> c.value(1) }
                    r.cell(2)
                    r.cell(3)
                }
            }
            w.sheet("One") { s ->
                s.freeze(1, 1)
                s.row { r -> r.cell("First Row") }

                s.row { r ->
                    r.cell("AC") { c -> c.value("AC") }
                    r.cell("BE") { c -> c.value("BE") }
                }

                s.row { r ->
                    r.style { st ->
                        st.align(Keywords.center, Keywords.left)
                        st.border { b ->
                            b.color("#abcdef")
                            b.style(Keywords.dashDotDot)
                        }
                        st.border(Keywords.right) { b -> b.color("#00ff00") }
                    }
                    r.cell { c ->
                        c.value("Hello")
                        c.name("Salutation")
                        c.width(Keywords.auto)
                    }
                    r.cell { c ->
                        c.style("bold")
                        c.value("World")
                        c.comment { comm -> comm.text("This cell has some fancy fg/bg").author("musketyr") }
                        c.width(50.0)
                    }
                    r.cell { c ->
                        c.style { st ->
                            st.format("d.m.y")
                            st.align(Keywords.center, Keywords.center)
                        }
                        c.value(today)
                        c.comment("This is a date!")
                        c.colspan(5)
                        c.rowspan(2)
                    }
                }
            }
            w.sheet("Links") { s ->
                s.freeze(1, 0)
                s.row { r ->
                    r.cell { c ->
                        c.value("Document (and a very long text)")
                        c.link(Keywords.to).name("Salutation")
                        c.width(Keywords.auto)
                    }
                    r.cell { c ->
                        c.value("File")
                        c.link(Keywords.to).file("text.txt")
                    }
                    r.cell { c ->
                        c.value("URL")
                        c.link(Keywords.to).url("https://www.google.com")
                    }
                    r.cell { c ->
                        c.value("Mail (plain)")
                        c.link(Keywords.to).email("vladimir@orany.cz")
                    }
                    r.cell { c ->
                        c.value("Mail (with subject)")
                        val email = LinkedHashMap<String, String>()
                        email.put("subject", "Testing Excel Builder")
                        email.put("body", "It is really great tools")
                        c.link(Keywords.to).email(email, "vladimir@orany.cz")
                    }
                }
            }
            w.sheet("Groups") { s ->
                s.row { r ->
                    r.cell("Headline 1")
                    r.group { g ->
                        g.cell { c ->
                            c.value("Headline 2")
                            c.style { st ->
                                st.foreground(aquamarine)
                                st.fill(Keywords.solidForeground)
                            }
                        }
                        g.cell("Headline 3")
                        g.collapse { clps ->
                            clps.cell("Headline 4")
                            clps.cell("Headline 5")
                        }
                        g.cell("Headline 6")
                    }
                }
                s.group { g ->
                    g.row { r -> r.cell("Some stuff") }
                    g.collapse { clps ->
                        clps.row { r -> r.cell("Something") }
                        clps.row { r -> r.cell("Something other") }
                    }
                    s.row { r -> r.cell("Other stuff") }
                }
            }
            w.sheet("Fills") { s ->
                for (foregroundFill in ForegroundFill.values()) {
                    s.row { r ->
                        r.cell { c ->
                            c.width(Keywords.auto)
                            c.value(foregroundFill.name)
                        }
                        r.cell { c ->
                            c.style { st ->
                                st.background("#FF8C00") // darkOrange
                                st.foreground(brown)
                                st.fill(foregroundFill)
                            }
                        }
                    }
                }
            }
            w.sheet("Formula") { s ->
                s.row { r ->
                    r.cell { c ->
                        c.value(10)
                        c.name("_Cell10")
                    }
                    r.cell { c ->
                        c.value(20)
                        c.name("_Cell20")
                    }
                    r.cell { c -> c.formula("SUM(#{_Cell10}:#{_Cell20})") }
                }
            }
            // create sheet with same name, it should query the
            w.sheet("Formula") { s ->
                s.row { r ->
                    r.cell { c ->
                        c.value(30)
                        c.name("_Cell30")
                    }
                    r.cell { c ->
                        c.value(40)
                        c.name("_Cell40")
                    }
                    r.cell { c -> c.formula("SUM(#{_Cell30}:#{_Cell40})") }
                }
            }
            w.sheet("Traversal") { s ->
                s.row { r ->
                    r.cell("A")
                    r.cell("B")
                    r.cell("C")
                }
                s.row { r ->
                    r.cell("D")
                    r.cell { c ->
                        c.value("E")
                        c.colspan(2)
                    }
                    r.cell("F")
                }
                s.row { r ->
                    r.cell("G")
                    r.cell { c ->
                        c.value("H")
                        c.rowspan(2)
                    }
                    r.cell("I")
                }
                s.row(5) { r -> r.cell("B") { c -> c.value("J") } }
            }
            w.sheet("Border") { s ->
                s.row { r ->
                    r.style("borders")
                    r.cell { c ->
                        c.value(1)
                        c.colspan(2)
                    }
                    r.cell { c -> c.value(2) }
                }
                s.row { r ->
                    r.style("borders")
                    r.cell { c ->
                        c.value(1)
                        c.colspan(2)
                    }
                    r.cell { c -> c.value(2) }

                }
            }
            w.sheet("Zero") { s -> s.row { r -> r.cell { c -> c.value(0) } } }
            w.sheet("Filtered") { s ->
                s.filter(Keywords.auto)
                s.row { r ->
                    r.cell("Name")
                    r.cell("Profession")
                }
                s.row { r ->
                    r.cell("Donald")
                    r.cell("Sailor")
                }
                s.row { r ->
                    r.cell("Bob")
                    r.cell("Builder")
                }
            }

            w.sheet("Styles") { s ->
                s.row { r ->
                    r.style("nonefg")
                    r.cell { c ->
                        c.value("NONE")
                        c.rowspan(3)
                        c.style("centered")
                    }
                }
                s.row(4) { r ->
                    r.style("redfg")
                    r.cell { c ->
                        c.value("RED")
                        c.rowspan(3)
                        c.style("centered")
                    }
                }
                s.row(7) { r ->
                    r.style("greenfg")
                    r.cell { c ->
                        c.value("GREEN")
                        c.rowspan(3)
                        c.styles("centered", "bold")
                    }
                }
                s.row(10) { r ->
                    r.styles("bluefg", "borders")
                    r.cell { c ->
                        c.value("BLUE")
                        c.rowspan(3)
                        c.styles("centered", "bold")
                    }
                }
            }

            w.sheet("Dimensions") { s ->
                s.row { r ->
                    r.cell { c ->
                        c.value("cm")
                        c.height(1.0).cm()
                        c.width(1.0).cm()
                    }
                }

                s.row { r ->
                    r.cell("B") { c ->
                        c.value("inches")
                        c.width(1.0).inch()
                        c.height(1.0).inch()
                    }
                }
                s.row { r ->
                    r.cell("C") { c ->
                        c.value("points")
                        c.width(10.0)
                        c.height(50.0)
                    }
                }
            }

            w.sheet("Custom Page") { s ->
                s.page { page ->
                    page.paper(Keywords.A5)
                    page.fit(Keywords.width).to(1)
                    page.orientation(Keywords.landscape)
                }
                s.row { r -> r.cell("A5 Landcapse") }
            }
            w.sheet("Broken row styles"
            ) { s -> s.row { r -> r.styles("bold", "redfg").cell { c -> c.value("bold and RED") } } }
            w.sheet("Hidden"
            ) { s -> s.state(Keywords.hidden) }
            w.sheet("Very hidden"
            ) { s -> s.state(Keywords.veryHidden) }
        }
    }

    private fun printMap(c: CellDefinition, entry: Entry<String, String>) {
        c.text(entry.key) { f -> f.style(Keywords.bold) }
        c.text(": ")
        c.text(entry.value)
        c.text("\n")
    }

}
