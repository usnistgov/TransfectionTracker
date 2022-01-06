# Transfection Tracker README

> Disclaimer: This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY that it will perform flawlessly or be
fit for a particular purpose. The use or mention by NIST of commercial products does not imply endorsement or indication that they are the
only, or best, products.

## Purpose of this program

This program allows collection of information about gene editing
operations and activities performed on transfected cells. It allows
tracking of what is done and when to individual cell samples over time.

## Instructions for using TransfectionTracker

**TransfectionTracker_v1221.xlsm**

## Table of contents

* [GENERAL OVERVIEW](#overview)
* [DATA ENTRY: STEP-BY-STEP](#dataentry)
* [FEATURES OVERVIEW](#features)
1)  [The **Calendar** worksheet](#calWksheet)
2)  [Navigating the **Calendar** worksheet](#navCal)
3)  [Using the *\#* and *Plate* columns](#plate)
4)  [***Activity*** *Discontinue*](#discontinue)
5)  [***Activity** Thaw*](#thaw)
6)  [***Activity** Sort*](#sort)
7)  [Other worksheets](#other)
8)  [Controlling and making changes to the worksheets](#controlWksheet)
9)  [**Do NOT...**](#donot)

<a name="overview"></a>
## GENERAL OVERVIEW: 

1)  This program was created in Microsoft Excel for Microsoft 365 MSO
    (16.0.13127.21490) 32-bit.

2)  The program is composed of interrelated worksheets that are accessed
    by named tabs that run along the bottom of the screen. Worksheets
    include **Calendar**, **Metadata Template**, **Cell Samples**,
    **ActivityList**, **Data Validation Criteria**, and
    **Documentation**.

3)  In this document, **Bold** lettering indicates the name of a
    worksheet. ***Bold Italics*** indicates an input available on a
    worksheet such as ***Activity*** and ***Cell Sample Name*** on the
    **Calendar** or ***Data*** from the toolbar at the top of the page.
    *Italics* indicates a choice of input often from a drop-down list.
    `Highlighted Text` indicates examples of free text.

4)  Two versions of the program are provided. One is populated with
    existing data so the user can see examples of what data are
    collected, how data are tabulated, and examples of reports created
    by querying the data. The version indicated by "\_clean" has no
    saved data. The user can begin to populate this version with new or
    test data, or they can modify it to suit their own data needs.

5)  Refer to the **Documentation** worksheet within the program for
    additional specific information.

<a name="dataentry"></a>
## DATA ENTRY: STEP-BY-STEP: 

1)  On opening the program, in response to dialog boxes, Enable Editing,
    Enable Content. If you intend to add data, answer NO if you see a
    question to open in ReadOnly mode.

2)  From the **Calendar** worksheet: Type in a date or use blue arrows
    on upper left of the **Calendar** to navigate.

    a.  From ***Activity*** for that date, choose *Transfect.* Use
        button at upper right of that day to ***Save***. A ***Metadata
        Sheet Selection*** box will appear; use the drop-down list to
        choose *Metadata Template* (subsequently, the new metadata sheet
        that is about to be created will be a choice option). Click
        ***Proceed***. A message that the data have been saved will
        appear; click *OK*. The **Metadata Template** will open. Cell D5
        will be prepopulated with the selected date.

        i.  Provide the minimum amount of information by making a
            selection from the drop-down lists in cells D11, G11, and
            D12. You will see error messages in cells L11 and L12 if the
            ***`Original plasmid backbone information`***
            (box to the right) contains data that are inconsistent with
            the information in D11, G11 and D12. An error here will not
            prevent proceeding. A transfection designator will be
            created in cell D44 with this information and the worksheet
            will be renamed.

3)  Navigate back to the **Calendar** worksheet. The cell sample name
    that corresponds to the transfection designator (and which is the
    name of the new metadata worksheet) now appears on the **Calendar**
    worksheet in the ***Cell Sample Name*** column corresponding to the
    ***Activity*** *Transfect*. Press the ***Save*** button again. The
    new transfection designator now appears in the drop-down list under
    ***Cell Sample Name***.

4)  After entering data on the **Calendar**, press the ***Save*** button
    before navigating away from the **Calendar** page.

<a name="features"></a>
## FEATURES OVERVIEW:

<a name="calWksheet"></a>

1)  The **Calendar** worksheet allows entry of activities performed on
    each sample for any date.

<a name="navCal"></a>

2)  In the **Calendar** worksheet, navigate to any date desired using
    the large BLUE arrows in the upper left of the Calendar window, or
    use the calendar icon to enter a date. In the box for the desired
    date, use drop-down arrows under headings ***Cell Sample Name*** and
    ***Activity*** to show lists of cell samples and activities to
    select from. The names of all cell samples that have been created
    will appear in the drop-down list and in the **Cell Samples**
    worksheet, which is updated when new samples are created.

    a.  The ***Activity,*** *Transfect,* provides a template for
        collecting metadata about a new transfection. The user can
        choose the **Metadata Template** worksheet or a previously
        created transfection worksheet to modify. As the data for the
        new transfection are entered in the worksheet, the name of the
        worksheet is updated to reflect the date (which is automatically
        entered when the ***Activity,*** *Transfect,* is initiated from
        the **Calendar**), and other data that the user inputs about
        that specific transfection including the gene being modified,
        the fluorescent protein sequence being used and a designation
        for the guide RNA. A transfection designator (line 44 on the
        worksheet) is assigned by concatenating this information and the
        worksheet is automatically renamed with the transfection
        designator, which is the new cell sample name. Thus for each
        transfection, a worksheet named with the date and other
        information about the transfection is created in the workbook.
        An electronic folder named with the transfection designator can
        be automatically created within the operating system by
        providing the url of a network computer in cell A6 of the
        Documentation worksheet.\*\*

    b.  When the new transfection worksheet and the corresponding
        transfection designator are created, the new cell sample name is
        added to the **Cell Samples** worksheet and becomes available in
        the drop-down menu in the **Calendar** worksheet under ***Cell
        Sample Name.*** The newest addition to the **Cell Samples**
        worksheet will be at the top of the ***Cell Sample Name***
        drop-down list.

    c.  In the use case shown here, each cell sample can be chosen on
        any day of the **Calendar** for an activity such as feeding,
        imaging, passaging, etc. (In this use case, the results of
        imaging measurements were used to determine whether or not there
        was a fluorescent transfected clone in the well plate.)

    d.  When a clone is identified, it is given a new name by choosing
        the button to the right of the date labeled **Designate New
        Clone**. The user is presented with a drop-down menu containing
        existing cell sample names from which they make a selection.
        This is followed by a free text window to assign a *Clone ID*,
        which is generally a position on a multi-well plate. The *Clone
        ID* will be appended to the original cell sample name and the
        newly created name will be added to the **Cell Samples**
        worksheet and appear in the ***Cell Sample Name*** drop-down
        list on the **Calendar**. A new subfolder designated with the
        name of the new clone can be created in the folder that was
        named for the transfection designator.\*\*

    e.  Worksheets can be temporarily hidden with the *Hide/Unhide*
        option by right clicking on the worksheet tab. This will reduce
        the number of worksheets visible (but accessible) in the
        workbook.

<a name="plate"></a>

3)  The columns labeled ***\#*** and ***Plate*** can be used to record
    the \# of wells used and the total number of wells in cell culture
    plate for activities such as *Passage*, *Sort*, and *Freeze*.
    Exceptions to this rule for entries in the ***Plate*** column:

    a.  `#` of `10cm` plates (type in 10cm)

    b.  `#` of `vials` for freezing (type in
        the word vial)

    c.  `#` of flasks of size (`T125`) (type in
        T125 or other flask designation)

<a name="discontinue"></a>

4)  Select ***Activity*** *Discontinue* when no versions of that clone
    are being carried further because of loss of fluorescence or other
    reason. Add reason to the ***Notes*** column as free text.

<a name="thaw"></a>

5)  The ***Activity** Thaw* applies only to cells from a bank previously
    subjected to ***Activity** Freeze*. If you take a sample from the
    bank, indicate ***Activity** Thaw*. This might be followed by
    ***Activity*** *ExtractDNA*.

    a.  A *Thaw* event is counted as incrementing passage number.

<a name="sort"></a>

6)  ***Activity** Sort* will append the **Cell Samples** worksheet with
    a new cell sample name to indicate sort number (\_Sort#). This new
    name will appear in the ***Cell Sample Name*** drop-down list on the
    **Calendar**. This action can also initiate the creation of a new
    folder with the appended name in the appropriate Transfection folder
    if that feature is enabled.\*\*

    a.  Details of ***Activity** Sort* can be captured as free text in
        adjacent ***Notes*** column.

    b.  The sort number automatically increments when a previously
        sorted sample is sorted again.

    c.  The ***Activity** Sort* is counted as an increment to passage
        number.

<a name="other"></a>

7)  Other worksheets:

    a.  Existing cell sample names are listed in worksheet **Cell
        Samples**. All activities are listed in worksheet **Activity
        List**.

    b.  Data collected on the **Calendar** worksheet are organized in
        tabular form in worksheet **Data**. These data can be in the
        form of a data range or designated as a table by highlighting
        all cells and selecting from the toolbar ***Insert*** / *Tables*
        / *Table*.

    c.  Examples of three **Transfection Metadata** worksheets (ex.
        **20200120mChOCT4sg2**) that were created for clones resulting
        from three transfections.

    d.  \*\*The **Documentation** worksheet contains instructions to
        direct the automatic creation of folders and files with human
        readable names on a network drive. By default, the lines of code
        that enable that function are comment lines.

    e.  Summary data can be created by queries. Examples of queries are
        **TransfectionsReport**, **DiscontinuedReport**,
        **FreezeReport**, **ExtractDNAReport**,
        **20200113mChOCT4sg2Report** and **Passage#s...** worksheets. As
        new data are added to the **Data** worksheet (such as from new
        entries on the **Calendar**, these **...Reports** worksheets can
        be updated from the toolbar by *Data* / *Refresh All*.

<a name="controlWksheet"></a>

8)  Controlling and making changes to the worksheets: (N.B.: These
    tips are meant to guide a user, not to serve as an Excel tutorial.)

    a.  Whether you are using TransfectionTracker_v1221.xlsm or
        TransfectionTracker_v1221.clean.xlsm, consider saving a renamed
        version that you can write over without changing the original
        file.

    b.  The **Metadata Template** has been designed for a particular use
        case, i.e., the collecting of information about clonal cell
        lines that are transfected by electroporation and identified and
        isolated by their fluorescence signal. The **Metadata Template**
        is created to capture some of the variations on a general
        protocol that are likely to occur and/or are being specifically
        tested. The metadata capture approach is designed to make it
        easy to record relevant specific details about the transfection
        and clonal isolation. Other use cases could require a
        modification or redesign of the template. This is relatively
        easy to do even without making changes to the VBA code by
        leaving in place cells that are specifically referred to in the
        code, namely D5 (date of transfection) and D44 (the transfection
        designator). If those positions are maintained, a new template
        could be created and substituted for the **Metadata Template**
        presented here. An example of modification of the template is
        shown by the differences between the metadata worksheet for
        transfection **20200120mChOCT4sg2** and the **Metadata
        Template.** The former was modified to add additional options on
        line 33 of the **Metadata Template** regarding the vessel for
        the transfection reaction.

        i.  If a different naming scheme is desired, one can change the
            concatenation argument in cell D44 of the **Metadata
            Template**. Unprotect the worksheet as described in 8d
            below. Cell D44 currently references the values of three
            cells, D5 (the date), H11 (the fluorescent protein) and D12
            (the guide RNA designation). Choose different cells to
            include in the concatenation instead, or create new cells
            with the desired new information, and reference those values
            in D44.

        ii. The **Data Validation criteria** worksheet can be edited to
            add options to existing drop-down lists. With your cursor on
            the cell that you want to add drop-down options to, choose
            ***Data*** on the toolbar and the *Data Validation* option
            under *Data Tools*. In the pop-up box, choose *List*, and as
            the *Source*, navigate to the **Data Validation** worksheet
            and highlight the cells that you want the user to be able to
            select. If you add terms to one of the lists, you will have
            to expand the size of the list by dragging your cursor to
            include the added options. You can always eliminate or
            substitute terms in these lists. If desired, you can protect
            this sheet to keep anyone else from making changes.

        iii. New drop-down lists can be added to the **Data Validation
             criteria** worksheet and can be selected from another
             worksheet as described above.

        iv. Drop-down lists can be filtered by highlighting the
            contents, selecting *Data* from the toolbar, and the
            *Filter* option from *Sort & Filter*.

    c.  Simple queries (the results of which are shown in the worksheets
        named **....Report**) can be generated by designating the
        **Data** worksheet contents as a table, highlighting the data,
        and selecting from the toolbar ***Data*** / *Get & Transform
        Data* / *From Table/Range* / *Table* (icon). A *Power Query
        Editor* dialog box appears, and data of interest can be selected
        from the ***Activities*** column. A new worksheet with the
        selected data is created which can be renamed. As new data are
        added to the **Data** worksheet (such as from new entries on the
        **Calendar**, these **...Reports** worksheets can be
        automatically updated from the toolbar by *Data* / *Refresh
        All*. A similar approach was used to generate the ...**Report**
        worksheets for each transfection by selecting on ***Cell Sample
        Name*** in the *Power Query Editor* dialog box. Further queries
        of these selected data were used to generate the
        **Passage#s...** worksheets. (N.B. Do not assign the **Cell
        Samples** worksheet as a table, since it needs to operate as a
        data range.)

    d.  Protecting and unprotecting workbooks, worksheets, and cells
        provides control by restricting worksheet cell entries while
        allowing intentional changes to be made. The protected status of
        the workbook and worksheets can be seen by choosing ***File*** /
        *Info*. Five worksheets in the current workbook are protected:
        **Calendar**, **Metadata Template** and metadata worksheets for
        three transfections. Protections can be removed from the
        ***File** / Info* page by pressing `*Unprotect*` or
        from the worksheet at the tool bar by choosing ***Review*** /
        *Unprotect Sheet*, by using the password `aplant`.
        From a worksheet, it is possible to designate, through
        ***Home*** / *Cells* / *Format* function, which cells are to be
        protected and which are to be available for user input. In the
        **Metadata Template** worksheet, the orange-colored cells are
        not protected so the user can enter or select values for those
        cells.

    e.  If entries are removed from the **Data** worksheet, they will be
        removed from the **Calendar** worksheet as well when the
        calendar date is scrolled.

    f.  The activity ***Transfect*** could be renamed to accommodate a
        different kind of experiment or to use this tracking system for
        routine handling of multiple cell samples. For example,
        ***Transfect*** could be substituted in the **ActivityList** for
        ***Start a new cell line*** and the **Metadata Template** could
        be modified as described in part 8b above. However, it would be
        necessary to be sure that the new activity name was substituted
        in all relevant lines in the VBA code.

<a name="donot"></a>

9)  **Do NOT:**

    a.  ...Type in a cell sample name or an activity into a cell on the
        **Calendar**. Choose only options from the drop-down lists. If
        additions to the drop-down lists are needed, add them in the
        **Data Validation** worksheet.

    b.  ...Insert a row, column or cell into the **Calendar**. If a row
        is added inadvertently, it may not have a drop-down arrow
        associated with it. If this happens, simply use a row for that
        date that does access a drop-down list.
