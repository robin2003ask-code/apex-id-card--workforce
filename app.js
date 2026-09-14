// =========================================
// ELEMENTS
// =========================================

const excelFile =
    document.getElementById("excelFile");

const photoFiles =
    document.getElementById("photoFiles");

const excelStatus =
    document.getElementById("excelStatus");

const photoStatus =
    document.getElementById("photoStatus");

const generateBtn =
    document.getElementById("generateBtn");

const pdfBtn =
    document.getElementById("pdfBtn");

const printBtn =
    document.getElementById("printBtn");

const clearBtn =
    document.getElementById("clearBtn");

const pageLayout =
    document.querySelector(".page-layout");

const cardsContainer =
    document.getElementById("cardsContainer");

const cardTemplate =
    document.querySelector(".card-template");

const companyName =
    document.getElementById("companyName");

const address1 =
    document.getElementById("address1");

const address2 =
    document.getElementById("address2");

const address3 =
    document.getElementById("address3");

const headerColor =
    document.getElementById("headerColor");


// =========================================
// VARIABLES
// =========================================

let employeeData = [];

let photos = {};


// =========================================
// DOWNLOAD EXCEL FORMAT
// =========================================

const downloadExcelBtn =
    document.getElementById("downloadExcelBtn");


if (downloadExcelBtn) {

    downloadExcelBtn.addEventListener(
        "click",
        () => {

            const data = [
                {
                    "EMP ID": "001",
                    "EMP NAME": "Sample Employee",
                    "DEPARTMENT": "HR",
                    "SUB DEPARTMENT": "ADMIN",
                    "DOJ": "01-01-2025",
                    "GENDER": "Male",
                    "DOB": "10-05-1995",
                    "RESIDENCE": "LabourShed 1"
                },

                {
                    "EMP ID": "002",
                    "EMP NAME": "Sample Employee 2",
                    "DEPARTMENT": "PRODUCTION",
                    "SUB DEPARTMENT": "SOLAR",
                    "DOJ": "15-02-2025",
                    "GENDER": "Female",
                    "DOB": "20-08-1996",
                    "RESIDENCE": "Dayscholar"
                }
            ];


            const worksheet =
                XLSX.utils.json_to_sheet(data);


            worksheet["!cols"] = [
                { wch: 15 },
                { wch: 25 },
                { wch: 20 },
                { wch: 22 },
                { wch: 15 },
                { wch: 12 },
                { wch: 15 },
                { wch: 18 }
            ];


            const workbook =
                XLSX.utils.book_new();


            XLSX.utils.book_append_sheet(
                workbook,
                worksheet,
                "Employees"
            );


            XLSX.writeFile(
                workbook,
                "Workforce IdCard Format.xlsx"
            );

        }
    );

}


// =========================================
// EXCEL UPLOAD
// =========================================

excelFile.addEventListener(
    "change",
    function (event) {

        const file =
            event.target.files[0];

        if (!file) return;


        const reader =
            new FileReader();


        reader.onload =
            function (e) {

                try {

                    const data =
                        new Uint8Array(
                            e.target.result
                        );


                    const workbook =
                        XLSX.read(
                            data,
                            {
                                type: "array"
                            }
                        );


                    const sheetName =
                        workbook.SheetNames[0];


                    const worksheet =
                        workbook.Sheets[
                        sheetName
                        ];


                    employeeData =
                        XLSX.utils.sheet_to_json(
                            worksheet,
                            {
                                defval: ""
                            }
                        );


                    excelStatus.textContent =
                        employeeData.length +
                        " employee records loaded";


                    console.log(
                        employeeData
                    );


                } catch (error) {

                    console.error(error);

                    excelStatus.textContent =
                        "Excel file error";

                }

            };


        reader.readAsArrayBuffer(file);

    }
);


// =========================================
// PHOTO UPLOAD
// =========================================

photoFiles.addEventListener(
    "change",
    function (event) {

        const files =
            event.target.files;


        photos = {};


        if (!files.length) {

            photoStatus.textContent =
                "No photos selected";

            return;

        }


        let loadedCount = 0;


        Array.from(files).forEach(
            function (file) {

                const reader =
                    new FileReader();


                reader.onload =
                    function (e) {

                        const fileName =
                            file.name
                                .split(".")
                                .slice(0, -1)
                                .join(".")
                                .trim()
                                .toLowerCase();


                        photos[fileName] =
                            e.target.result;


                        loadedCount++;


                        photoStatus.textContent =
                            loadedCount +
                            " photos selected";

                    };


                reader.readAsDataURL(file);

            }
        );

    }
);


// =========================================
// GET VALUE
// =========================================
// Case-insensitive Excel column matching

function getValue(
    employee,
    keys
) {

    for (const key of keys) {

        const foundKey =
            Object.keys(employee).find(
                function (excelKey) {

                    return (
                        excelKey
                            .trim()
                            .toLowerCase() ===
                        key
                            .trim()
                            .toLowerCase()
                    );

                }
            );


        if (
            foundKey &&
            employee[foundKey] !== undefined &&
            employee[foundKey] !== null &&
            employee[foundKey] !== ""
        ) {

            return employee[foundKey];

        }

    }


    return "";

}


// =========================================
// DATE FORMAT
// =========================================

function formatDate(value) {

    if (!value) {

        return "";

    }


    // Excel serial date

    if (
        typeof value === "number"
    ) {

        const date =
            XLSX.SSF.parse_date_code(
                value
            );


        if (date) {

            return (
                String(date.d)
                    .padStart(2, "0")
                +
                "-" +
                String(date.m)
                    .padStart(2, "0")
                +
                "-" +
                date.y
            );

        }

    }


    return value;

}


// =========================================
// GENERATE BUTTON
// =========================================

generateBtn.addEventListener(
    "click",
    function () {

        if (
            employeeData.length === 0
        ) {

            alert(
                "Please upload Excel file first."
            );

            return;

        }


        // Hide input page

        pageLayout.style.display =
            "none";


        // Show cards

        cardsContainer.classList.add(
            "active"
        );


        generateCards();

    }
);


// =========================================
// GENERATE A4 PAGES
// =========================================

function generateCards() {

    // Clear previous cards

    cardsContainer.innerHTML = "";


    let currentPage = null;


    employeeData.forEach(
        function (
            employee,
            index
        ) {


            // =================================
            // NEW A4 PAGE EVERY 8 CARDS
            // =================================

            if (
                index % 8 === 0
            ) {

                currentPage =
                    document.createElement(
                        "div"
                    );


                currentPage.className =
                    "a4-page";


                currentPage.dataset.page =
                    Math.floor(
                        index / 8
                    ) + 1;


                cardsContainer.appendChild(
                    currentPage
                );

            }


            // =================================
            // CLONE CARD
            // =================================

            const card =
                cardTemplate.cloneNode(
                    true
                );


            card.classList.remove(
                "card-template"
            );


            card.classList.add(
                "generated-card"
            );


            card.style.display =
                "block";


            // =================================
            // EMPLOYEE DATA
            // =================================

            const empId =
                getValue(
                    employee,
                    [
                        "EMP ID",
                        "EMP_ID",
                        "Employee ID",
                        "Employee_ID",
                        "ID"
                    ]
                );


            const empName =
                getValue(
                    employee,
                    [
                        "EMP NAME",
                        "EMP_NAME",
                        "Employee Name",
                        "Employee_Name",
                        "NAME"
                    ]
                );


            // =================================
            // DEPARTMENT
            // =================================

            const dept =
                getValue(
                    employee,
                    [
                        "DEPT",
                        "DEPARTMENT",
                        "Department",
                        "department"
                    ]
                );


            // =================================
            // SUB DEPARTMENT
            // =================================

            const subDept =
                getValue(
                    employee,
                    [
                        "SUB-DEPT",
                        "SUB DEPT",
                        "SUB_DEPT",
                        "SUB DEPARTMENT",
                        "Sub Department",
                        "sub department"
                    ]
                );


            const doj =
                getValue(
                    employee,
                    [
                        "DOJ",
                        "Date of Joining",
                        "Joining Date"
                    ]
                );


            const gender =
                getValue(
                    employee,
                    [
                        "GENDER",
                        "Gender"
                    ]
                );


            const dob =
                getValue(
                    employee,
                    [
                        "DOB",
                        "Date of Birth"
                    ]
                );


            const residence =
                getValue(
                    employee,
                    [
                        "RESIDENCE",
                        "Residence"
                    ]
                );


            // =================================
            // SET EMPLOYEE DATA
            // =================================

            const empIdElement =
                card.querySelector(
                    ".emp-id"
                );


            const empNameElement =
                card.querySelector(
                    ".emp-name"
                );


            const deptElement =
                card.querySelector(
                    ".dept"
                );


            const subDeptElement =
                card.querySelector(
                    ".sub-dept"
                );


            const dojElement =
                card.querySelector(
                    ".doj"
                );


            const genderElement =
                card.querySelector(
                    ".gender"
                );


            const dobElement =
                card.querySelector(
                    ".dob"
                );


            if (empIdElement) {

                empIdElement.textContent =
                    empId;

            }


            if (empNameElement) {

                empNameElement.textContent =
                    empName;

            }


            if (deptElement) {

                deptElement.textContent =
                    dept;

            }


            if (subDeptElement) {

                subDeptElement.textContent =
                    subDept;

            }


            if (dojElement) {

                dojElement.textContent =
                    formatDate(doj);

            }


            if (genderElement) {

                genderElement.textContent =
                    gender;

            }


            if (dobElement) {

                dobElement.textContent =
                    formatDate(dob);

            }


            // =================================
            // COMPANY DETAILS
            // =================================

            const cardCompany =
                card.querySelector(
                    ".company-name"
                );


            const cardAddress1 =
                card.querySelector(
                    ".address1"
                );


            const cardAddress2 =
                card.querySelector(
                    ".address2"
                );


            const cardAddress3 =
                card.querySelector(
                    ".address3"
                );


            if (cardCompany) {

                cardCompany.textContent =
                    companyName.value;

            }


            if (cardAddress1) {

                cardAddress1.textContent =
                    address1.value;

            }


            if (cardAddress2) {

                cardAddress2.textContent =
                    address2.value;

            }


            if (cardAddress3) {

                cardAddress3.textContent =
                    address3.value;

            }


            // =================================
            // HEADER COLOR
            // =================================

            const header =
                card.querySelector(
                    ".card-header"
                );


            if (header) {

                let residenceColor =
                    "";

                const residenceValue =
                    String(residence)
                        .trim()
                        .toLowerCase();

                if (residenceValue === "dayscholar") {

                    residenceColor =
                        "#C3D79A";

                } else if (
                    residenceValue === "labourshed 2"
                ) {

                    residenceColor =
                        "#CDC0DA";

                } else if (
                    residenceValue === "hosteller"
                ) {

                    residenceColor =
                        "#B7DDE8";

                } else if (
                    residenceValue === "labourshed 1"
                ) {

                    residenceColor =
                        "#f3f33d";

                }

                header.style.backgroundColor =
                    residenceColor;

            }


            // =================================
            // EMPLOYEE PHOTO
            // =================================

            const employeePhoto =
                card.querySelector(
                    ".employee-photo"
                );


            const photoPlaceholder =
                card.querySelector(
                    ".photo-placeholder"
                );


            const photoKey =
                String(empId)
                    .trim()
                    .toLowerCase();


            if (
                photos[photoKey]
            ) {

                employeePhoto.src =
                    photos[photoKey];


                employeePhoto.classList.add(
                    "loaded"
                );


                if (
                    photoPlaceholder
                ) {

                    photoPlaceholder.style.display =
                        "none";

                }

            } else {

                employeePhoto.style.display =
                    "none";


                if (
                    photoPlaceholder
                ) {

                    photoPlaceholder.style.display =
                        "flex";

                }

            }


            // =================================
            // ADD CARD TO A4 PAGE
            // =================================

            currentPage.appendChild(
                card
            );

        }
    );


    console.log(
        employeeData.length +
        " ID cards generated."
    );

}


// =========================================
// PDF GENERATION
// =========================================

pdfBtn.addEventListener(
    "click",
    async function () {

        const pages =
            document.querySelectorAll(
                ".a4-page"
            );


        if (
            pages.length === 0
        ) {

            alert(
                "Please generate ID cards first."
            );

            return;

        }


        const {
            jsPDF
        } = window.jspdf;


        const pdf =
            new jsPDF(
                {
                    orientation: "portrait",
                    unit: "mm",
                    format: "a4"
                }
            );


        for (
            let i = 0;
            i < pages.length;
            i++
        ) {

            if (i > 0) {

                pdf.addPage();

            }


            const canvas =
                await html2canvas(
                    pages[i],
                    {
                        scale: 2,

                        useCORS: true,

                        backgroundColor:
                            "#ffffff"
                    }
                );


            const image =
                canvas.toDataURL(
                    "image/png"
                );


            pdf.addImage(
                image,
                "PNG",
                0,
                0,
                210,
                297
            );

        }


        pdf.save(
            "Employee-ID-Cards-A4.pdf"
        );

    }
);


// =========================================
// PRINT
// =========================================

printBtn.addEventListener(
    "click",
    function () {

        const pages =
            document.querySelectorAll(
                ".a4-page"
            );


        if (
            pages.length === 0
        ) {

            alert(
                "Please generate ID cards first."
            );

            return;

        }


        window.print();

    }
);


// =========================================
// CLEAR
// =========================================

clearBtn.addEventListener(
    "click",
    function () {

        employeeData = [];

        photos = {};


        excelFile.value = "";

        photoFiles.value = "";


        excelStatus.textContent =
            "No Excel selected";


        photoStatus.textContent =
            "No photos selected";


        cardsContainer.innerHTML =
            "";


        cardsContainer.classList.remove(
            "active"
        );


        pageLayout.style.display =
            "block";

    }
);

