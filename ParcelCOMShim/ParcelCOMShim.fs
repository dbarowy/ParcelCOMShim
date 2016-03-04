namespace ParcelCOMShim
    open System
    open Microsoft.Office.Interop.Excel
    open Parcel

    type XLRange = Microsoft.Office.Interop.Excel.Range

    type COMRef(wb: Workbook,
                ws: Worksheet,
                range: XLRange,
                path: string,           // path excluding final separator and filename; option type because in-memory workbooks have no path
                workbook_name: string,
                worksheet_name: string,
                formula: string option,
                width: int,
                height: int) =
        let _wb = wb
        let _ws = ws
        let _r = range
        let _is_cell = width = 1 && height = 1
        let _width = width
        let _height = height
        let _path = path
        let _workbook_name = workbook_name
        let _worksheet_name = worksheet_name
        let _formula = formula
        let mutable _do_not_perturb = match formula with | Some(f) -> true | None -> false

        member self.Width = _width
        member self.Height = _height
        member self.Workbook = _wb
        member self.Worksheet = _ws
        member self.Range = _r
        member self.IsFormula = match _formula with | Some(f) -> true | None -> false
        member self.Formula = match _formula with
            | Some(f) -> f
            | None -> failwith "Not a formula reference."
        member self.IsCell = _is_cell
        member self.Path = _path
        member self.WorkbookName = _workbook_name
        member self.WorksheetName = _worksheet_name
        member self.DoNotPerturb
            with get() = _do_not_perturb
            and set(value) = _do_not_perturb <- value
        override self.GetHashCode() =
            let x = range.Column
            let y = range.Row
            AST.Address.addressHash x y _worksheet_name

    module Address =
        let GetCOMObject(addr: AST.Address, app: Application) : XLRange =
            let wb: Workbook = app.Workbooks.Item(addr.A1Workbook())
            let ws: Worksheet = wb.Worksheets.Item(addr.A1Worksheet()) :?> Worksheet
            let cell: XLRange = ws.Range(addr.A1Local())
            cell

        let AddressFromCOMObject(com: Microsoft.Office.Interop.Excel.Range, wb: Microsoft.Office.Interop.Excel.Workbook) : AST.Address =
            let wsname = com.Worksheet.Name
            let wbname = wb.Name
            let path = System.IO.Path.GetDirectoryName(wb.FullName)
            let addr = com.get_Address(true, true, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1, Type.Missing, Type.Missing)
            AST.Address.FromString(addr, wsname, wbname, path)

    module Range =
        // build a contiguous range from topleft and bottomright coordinates
        let private constructRange(app: Application)(tl: AST.Address)(br: AST.Address) : XLRange =
            let wb: Workbook = app.Workbooks.Item(tl.A1Workbook())
            let ws: Worksheet = app.Worksheets.Item(tl.A1Worksheet()) :?> Worksheet
            ws.Range(tl.A1Local(), br.A1Local())

        // build a discontiguous range from range components
        let GetCOMObject(rng: AST.Range, app: Application) : XLRange =
            let rngs = rng.Ranges() |> List.map (fun (tl: AST.Address, br: AST.Address) -> constructRange app tl br)
            List.reduce (fun acc r -> app.Union(acc, r)) rngs