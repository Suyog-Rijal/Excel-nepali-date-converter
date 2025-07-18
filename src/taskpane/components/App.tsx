import React, { FC, useEffect, useRef, useState } from "react";
import { Button, makeStyles } from "@fluentui/react-components";
import CustomDropdown from "./CustomDropdown";

const useStyles = makeStyles({
    app: {
        minHeight: "100vh",
        display: "flex",
        flexDirection: "column",
        alignItems: "center"
    },
    h1: {
        textAlign: "center",
        lineHeight: "2rem"
    }
});

const App: FC = () => {
    const BASEURL = "https://excel-nepali-date-converter-backend.onrender.com";
    const styles = useStyles();
    const [columns, setColumns] = useState<{ key: string; value: string }[]>([]);
    const [selectedColumn, setSelectedColumn] = useState<string | null>(null);
    const [operations] = useState<{ key: string; value: string }[]>([
        { key: "bs-to-ad", value: "BS to AD" },
        { key: "ad-to-bs", value: "AD to BS" }
    ]);
    const [selectedOperation, setSelectedOperation] = useState<string>(operations[0].key);
    const [dateRangeLock] = useState<{ key: string; value: string }[]>([
        { key: "auto", value: "Auto" },
        { key: "disable", value: "Disable" }
    ]);
    const [selectedDateRangeLock, setSelectedDateRangeLock] = useState<string>("auto");
    const [isMonitoring, setIsMonitoring] = useState(false);
    const currentYearRef = useRef<number | null>(null);
    const dateLookupRang = 10; // 10 years for date range lock

    const isMonitoringRef = useRef(isMonitoring);
    const selectedColumnRef = useRef(selectedColumn);
    const selectedOperationRef = useRef(selectedOperation);
    const selectedDateRangeLockRef = useRef(selectedDateRangeLock);

    useEffect(() => {
        isMonitoringRef.current = isMonitoring;
    }, [isMonitoring]);

    useEffect(() => {
        selectedColumnRef.current = selectedColumn;
    }, [selectedColumn]);

    useEffect(() => {
        selectedOperationRef.current = selectedOperation;
        if (selectedOperation === "bs-to-ad") {
            currentYearRef.current = new Date().getFullYear() + 56;
        }
        else if (selectedOperation === "ad-to-bs") {
            currentYearRef.current = new Date().getFullYear();
        } else {
            currentYearRef.current = null;
        }
    }, [selectedOperation]);

    useEffect(() => {
        selectedDateRangeLockRef.current = selectedDateRangeLock;
    }, [selectedDateRangeLock]);

    useEffect(() => {
        scanColumns().then();
    }, []);

    useEffect(() => {
        const handler = async (args: Excel.WorksheetChangedEventArgs) => {
            const fullAddr = args.address.split("!").pop()!;
            const changedAddr = fullAddr.split(":")[0];
            if (!selectedColumnRef.current || !changedAddr.startsWith(selectedColumnRef.current)) return;
            if (!isMonitoringRef.current) return;

            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const cell = sheet.getRange(changedAddr);
                cell.load("values");
                await context.sync();

                const newVal = String(cell.values[0][0] ?? "");

                await handleConversion(changedAddr, newVal);
            });
        };

        let eventHandler: any;
        Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            eventHandler = sheet.onChanged.add(handler);
            await context.sync();
        });

        return () => {
            Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.onChanged.remove(eventHandler);
                await context.sync();
            });
        };
    }, []);

    const scanColumns = async () => {
        const getColumnLetter = (colIndex: number): string => {
            let letter = "";
            while (colIndex >= 0) {
                letter = String.fromCharCode((colIndex % 26) + 65) + letter;
                colIndex = Math.floor(colIndex / 26) - 1;
            }
            return letter;
        };

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const usedRange = sheet.getUsedRange();
                usedRange.load(["columnCount", "rowCount", "values"]);
                await context.sync();

                const { columnCount, rowCount, values } = usedRange;
                const result: { key: string; value: string }[] = [];

                for (let col = 0; col < columnCount; col++) {
                    for (let row = 0; row < rowCount; row++) {
                        const cellValue = values[row][col];
                        if (cellValue !== null && cellValue !== "") {
                            const columnLetter = getColumnLetter(col);
                            result.push({
                                key: columnLetter,
                                value: `${String(cellValue)} (Col ${columnLetter})`
                            });
                            break;
                        }
                    }
                }

                setColumns(result);
            });
        } catch {
            setColumns([]);
            setSelectedColumn(null);
        }
    };

    const handleConversion = async (
        address: string,
        current: string) => {
        const data = parseDate(current);
        if (!data) {
            return;
        }
        const { day, month, year } = data;
        if (year > currentYearRef.current + dateLookupRang || year < currentYearRef.current - dateLookupRang) {
            return;
        }

        const url = `${BASEURL}/${selectedOperationRef.current}`;
        const payload = {
            year: String(year),
            month: String(month).padStart(2, "0"),
            day: String(day).padStart(2, "0")
        };

        try {
            const response = await fetch(url, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                console.error("Server responded with error:", response.status);
                return;
            }

            const result = await response.json();
            if (result.error) {
                console.error("Server error:", result.error);
                return;
            }

            const formattedDate = `${String(result.month).padStart(2, "0")}/${String(result.day).padStart(2, "0")}/${result.year}`;
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const cell = sheet.getRange(address);
                cell.values = [[formattedDate]];
                await context.sync();
            });
        } catch (err) {
            console.error("Network or server error:", err);
        }

    };

    const parseDate = (value: string): { day: number; month: number; year: number } | null => {
        const serial = Number(value);
        if (!isNaN(serial)) {
            const jsDate = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
            return {
                day: jsDate.getUTCDate(),
                month: jsDate.getUTCMonth() + 1,
                year: jsDate.getUTCFullYear()
            };
        }

        const datePattern = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
        if (datePattern.test(value)) {
            const [month, day, year] = value.split("/").map(Number);
            return { day, month, year };
        }

        return null;
    };

    const handleMonitoring = async () => {
        setIsMonitoring(prev => !prev);
    };

    return (
        <div className={styles.app} style={{ position: "relative" }}>
            <div style={{ position: "absolute", bottom: 0, right: 0, margin: "1rem" }}>
                <Button onClick={scanColumns}>⭮ Refresh</Button>
            </div>
            <h1 className={styles.h1}>Welcome to NDC</h1>
            <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
                <CustomDropdown
                    id="column-selector"
                    label="Select a Column"
                    placeholder="Select col containing dates"
                    disabled={isMonitoring}
                    value={selectedColumn}
                    options={columns}
                    onChange={(value) => setSelectedColumn(value)}
                />
                <CustomDropdown
                    id="operation-selector"
                    label="Select Conversion Operation"
                    value={selectedOperation}
                    options={operations}
                    onChange={(value) => setSelectedOperation(value)}
                />
                <CustomDropdown
                    id="date-range-lock"
                    label="Select Date Range Lock"
                    value={selectedDateRangeLock}
                    options={dateRangeLock}
                    onChange={(value) => setSelectedDateRangeLock(value)}
                />
            </div>
            {selectedColumn && (
                <Button
                    onClick={handleMonitoring}
                    style={{ marginTop: "1rem" }}
                    appearance={isMonitoring ? "primary" : "secondary"}
                >
                    {isMonitoring ? "Stop Conversion" : "Start Conversion"}
                </Button>
            )}
        </div>
    );
};

export default App;