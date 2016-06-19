namespace ParcelCOMShim
    open System
    open Microsoft.Office.Interop.Excel
    open System.Runtime.Serialization
    open Parcel

    module Address =
        let GetCOMObject(addr: AST.Address, app: Application) : Range =
            let wb: Workbook = app.Workbooks.Item(addr.A1Workbook())
            let ws: Worksheet = wb.Worksheets.Item(addr.A1Worksheet()) :?> Worksheet
            let cell: Range = ws.Range(addr.A1Local())
            cell

        let AddressFromCOMObject(com: Range, wb: Workbook) : AST.Address =
            let wsname = com.Worksheet.Name
            let wbname = wb.Name
            let path = System.IO.Path.GetDirectoryName(wb.FullName)
            let addr = com.get_Address(true, true, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1, Type.Missing, Type.Missing)
            AST.Address.FromString(addr, wsname, wbname, path)

    module Range =
        // build a contiguous range from topleft and bottomright coordinates
        let private constructRange(app: Application)(tl: AST.Address)(br: AST.Address) : Range =
            let wb: Workbook = app.Workbooks.Item(tl.A1Workbook())
            let ws: Worksheet = app.Worksheets.Item(tl.A1Worksheet()) :?> Worksheet
            ws.Range(tl.A1Local(), br.A1Local())

        // build a discontiguous range from range components
        let GetCOMObject(rng: AST.Range, app: Application) : Range =
            let rngs = rng.Ranges() |> List.map (fun (tl: AST.Address, br: AST.Address) -> constructRange app tl br)
            List.reduce (fun acc r -> app.Union(acc, r)) rngs

    [<AbstractClass>]
    type COMRef() =
        abstract Width : int
        abstract Height : int
        abstract Workbook : Workbook
        abstract Worksheet : Worksheet
        abstract Range : Range
        abstract IsFormula : bool
        abstract Formula : string
        abstract IsCell : bool
        abstract Path : string
        abstract WorkbookName : string
        abstract WorksheetName : string
        abstract DoNotPerturb : bool with get, set
        abstract DeserializationCellFixup: AST.Address -> Application -> COMRef
        abstract DeserializationRangeFixup: AST.Range -> Application -> COMRef

    [<Serializable>]
    type NonLocalComRef(path: string,           // path excluding final separator and filename; option type because in-memory workbooks have no path
                        workbook_name: string,
                        worksheet_name: string,
                        formula: string option) =
        inherit COMRef()
        let _path = path
        let _workbook_name = workbook_name
        let _worksheet_name = worksheet_name
        let _formula = formula
        let mutable _do_not_perturb = match formula with | Some(f) -> true | None -> false

        override self.Width = raise (NotImplementedException())
        override self.Height = raise (NotImplementedException())
        override self.Workbook = raise (NotImplementedException())
        override self.Worksheet = raise (NotImplementedException())
        override self.Range = raise (NotImplementedException())
        override self.IsFormula = match _formula with | Some(f) -> true | None -> false
        override self.Formula =
            match _formula with
            | Some(f) -> f
            | None -> failwith "Not a formula reference."
        override self.IsCell = raise (NotImplementedException())
        override self.Path = _path
        override self.WorkbookName = _workbook_name
        override self.WorksheetName = _worksheet_name
        override self.DoNotPerturb
            with get() = _do_not_perturb
            and set(value) = _do_not_perturb <- value
        override self.DeserializationCellFixup(addr: AST.Address)(app: Application) : COMRef = self :> COMRef
        override self.DeserializationRangeFixup(rng: AST.Range)(app: Application) : COMRef = self :> COMRef

    [<Serializable>]
    type LocalCOMRef(wb: Workbook,
                     ws: Worksheet,
                     range: Range,
                     path: string,           // path excluding final separator and filename; option type because in-memory workbooks have no path
                     workbook_name: string,
                     worksheet_name: string,
                     formula: string option,
                     width: int,
                     height: int) =
        inherit COMRef()
        [<NonSerialized>]
        let _wb = wb
        [<NonSerialized>]
        let _ws = ws
        [<NonSerialized>]
        let _r = range
        let _is_cell = width = 1 && height = 1
        let _width = width
        let _height = height
        let _path = path
        let _workbook_name = workbook_name
        let _worksheet_name = worksheet_name
        let _formula = formula
        let mutable _do_not_perturb = match formula with | Some(f) -> true | None -> false

        override self.Width = _width
        override self.Height = _height
        override self.Workbook = _wb
        override self.Worksheet = _ws
        override self.Range = _r
        override self.IsFormula = match _formula with | Some(f) -> true | None -> false
        override self.Formula =
            match _formula with
            | Some(f) -> f
            | None -> failwith "Not a formula reference."
        override self.IsCell = _is_cell
        override self.Path = _path
        override self.WorkbookName = _workbook_name
        override self.WorksheetName = _worksheet_name
        override self.DoNotPerturb
            with get() = _do_not_perturb
            and set(value) = _do_not_perturb <- value
        override self.DeserializationCellFixup(addr: AST.Address)(app: Application) : COMRef =
            let cell = Address.GetCOMObject(addr, app);
            let wb = cell.Worksheet.Parent :?> Workbook
            let ws = cell.Worksheet
            // sanity check
            assert(wb.Name = _workbook_name && ws.Name = _worksheet_name)
            LocalCOMRef(wb, ws, cell, _path, _workbook_name, _worksheet_name, _formula, _width, _height) :> COMRef
        override self.DeserializationRangeFixup(rng: AST.Range)(app: Application) : COMRef =
            let r = Range.GetCOMObject(rng, app);
            let wb = r.Worksheet.Parent :?> Workbook
            let ws = r.Worksheet
            // sanity check
            assert(wb.Name = _workbook_name && ws.Name = _worksheet_name)
            LocalCOMRef(wb, ws, r, _path, _workbook_name, _worksheet_name, _formula, _width, _height) :> COMRef