const body = document.querySelector("body");



//login Select

const loginSection = document.querySelector(".login-section");

const loginHeading = document.querySelector(".login-heading");

const userName = document.querySelector(".userNameInput");
const password = document.querySelector(".userPassword");

const loginBtn = document.querySelector(".loginBtn").addEventListener("click", () => {

    event.preventDefault();


    const userName = document.querySelector(".userNameInput").value;
    const password = document.querySelector(".userPassword").value;


    if (userName !== "" && password !== "") {


        if (userName === "Robin") {

            if (password === "W14011" || password === "w14011") {
                loginHeading.style = "transition: 5s"
                loginHeading.innerHTML = "Success";

                setTimeout(() => {
                    loginSection.style.display = "none";
                    body.style = "background-image: url(asset/Background.png)";
                    navBar.style.display = "grid"
                }, 1000)
            }

            else{

                loginHeading.innerHTML = "Invaid Password";
                userName.value = "";
            }

        }

        else{

            loginHeading.innerHTML = "Invalid UserName";                        
        }

    }

    else {

        alert("Enter Your UserName & Password");
    }


})



//nav
const navBar = document.querySelector("nav");

//Home Section
const home = document.querySelector(".home").addEventListener("click", () => {

    idCardPage.style = "display: none";
    btnContainer.style = "display: none";

})

//menu selection
const menuIdcardPage = document.querySelector(".menuIdPage").addEventListener("click", () => {

    idCardPage.style = "display: grid";
    btnContainer.style = "display: flex";
    home.style = "background: red";

})

const menuKbReducerPage = document.querySelector(".menuImgReducerPage").addEventListener("click", () => {

    idCardPage.style = "display: none";
    btnContainer.style = "display: none";
})

// About Section
const about = document.querySelector(".about")
about.addEventListener("click", () => {

    idCardPage.style = "display: none";
    btnContainer.style = "display: none";
})

// LoggOff Section
const logOff = document.querySelector(".logoff").addEventListener("click", () => {

    loginSection.style = "display: flex";
    body.style = "background-image: url(asset/LoginBackground.png)";

    userName.value = "";
    password.value = "";
    loginHeading.innerHTML = "Hello!"

    navBar.style = "display: none";

    idCardPage.style = "display: none";
    btnContainer.style = "display: none";

})

//idCard Section
const idCardSection = document.querySelector(".idCard-section");
// idCard page
const idCardPage = document.querySelector(".idCardPage");

//Id Card other details  Section

const menuInside = document.querySelector(".menu-inside");

const openMenu = document.querySelector(".openMenuIcon");

openMenu.addEventListener("click", () => {

    menuInside.style = "transition: all 1s ease; opacity: 100%";

    openMenu.style = "display: none";

})

const colseMenu = document.querySelector(".closeMenuIcon")

colseMenu.addEventListener("click", () => {

    menuInside.style = "transition: all 1s ease; opacity: 0%;";

    openMenu.style = "display: block";

})


// =========================================
// ELEMENTS
// =========================================

const excelFile = document.getElementById("excelFile");
const photoFiles = document.getElementById("photoFiles");

const excelStatus = document.getElementById("excelStatus");
const photoStatus = document.getElementById("photoStatus");

const btnContainer = document.getElementById("btn-container");

const generateBtn = document.getElementById("generateBtn");
const pdfBtn = document.getElementById("pdfBtn");
const clearBtn = document.getElementById("clearBtn");

const cardsContainer = document.getElementById("cardsContainer");
const cardTemplate = document.querySelector(".card-template");

const companyName = document.getElementById("companyName");

const address1 = document.getElementById("address1");
const address2 = document.getElementById("address2");
const address3 = document.getElementById("address3");

const headerColor = document.getElementById("headerColor");

const idCardPage2 = document.querySelector(".idCardPage2");


// =========================================
// VARIABLES
// =========================================

let employeeData = [];
let photos = {};


// =========================================
// DOWNLOAD EXCEL FORMAT
// =========================================

const downloadExcelBtn = document.getElementById("downloadExcelBtn");

if (downloadExcelBtn) {
    downloadExcelBtn.addEventListener("click", () => {

        const data = [
            {
                "EMP ID": "001",
                "EMP NAME": "SAMPLE EMPLOYEE 1",
                "DEPARTMENT": "HR",
                "SUB DEPARTMENT": "ADMIN",
                "DOJ": "01-01-2025",
                "GENDER": "Male",
                "DOB": "10-05-1995",
                "RESIDENCE": "LABOUR SHED-I"
            },
            {
                "EMP ID": "002",
                "EMP NAME": "SAMPLE EMPLOYEE 2",
                "DEPARTMENT": "PRODUCTION",
                "SUB DEPARTMENT": "SOLAR",
                "DOJ": "15-02-2025",
                "GENDER": "Female",
                "DOB": "20-08-1996",
                "RESIDENCE": "DAYSCHOLAR"
            }
        ];

        const worksheet = XLSX.utils.json_to_sheet(data);

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

        const workbook = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(
            workbook,
            worksheet,
            "Employees"
        );

        XLSX.writeFile(
            workbook,
            "Workforce IdCard Format.xlsx"
        );
    });
}


// =========================================
// EXCEL UPLOAD
// =========================================

excelFile.addEventListener("change", function (event) {

    const file = event.target.files[0];

    if (!file) return;

    const reader = new FileReader();

    reader.onload = function (e) {

        try {
            const data = new Uint8Array(e.target.result);

            const workbook = XLSX.read(data, {
                type: "array"
            });

            const sheetName = workbook.SheetNames[0];

            const worksheet = workbook.Sheets[sheetName];

            employeeData = XLSX.utils.sheet_to_json(
                worksheet,
                {
                    defval: ""
                }
            );

            excelStatus.textContent =
                employeeData.length + " employee records loaded";

            console.log(employeeData);

        } catch (error) {
            console.error(error);

            excelStatus.textContent = "Excel file error";
        }
    };

    reader.readAsArrayBuffer(file);
});


// =========================================
// PHOTO UPLOAD
// =========================================

photoFiles.addEventListener("change", function (event) {

    const files = event.target.files;

    photos = {};

    if (!files.length) {
        photoStatus.textContent = "No photos selected";
        return;
    }

    let loadedCount = 0;

    Array.from(files).forEach(function (file) {

        const reader = new FileReader();

        reader.onload = function (e) {

            const fileName = file.name
                .split(".")
                .slice(0, -1)
                .join(".")
                .trim()
                .toLowerCase();

            photos[fileName] = e.target.result;

            loadedCount++;

            photoStatus.textContent =
                loadedCount + " photos selected";
        };

        reader.readAsDataURL(file);
    });
});


// =========================================
// GET VALUE
// =========================================

function getValue(employee, keys) {

    for (const key of keys) {

        const foundKey = Object.keys(employee).find(
            function (excelKey) {
                return (
                    excelKey.trim().toLowerCase() ===
                    key.trim().toLowerCase()
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
    if (typeof value === "number") {

        const date = XLSX.SSF.parse_date_code(value);

        if (date) {
            return (
                String(date.d).padStart(2, "0") +
                "-" +
                String(date.m).padStart(2, "0") +
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

generateBtn.addEventListener("click", function () {

    if (employeeData.length === 0) {
        alert("Please upload Excel file first.");
        return;
    }

    // Hide input page
    generateBtn.style = "display: none";
    pdfBtn.style.display = "block";

    idCardPage.style.display = "none";
    idCardPage2.style.display = "flex";

    navBar.style.display = "none";
    btnContainer.style = "display: flex;bottom: 0px";

    // Show cards
    cardsContainer.classList.add("active");

    generateCards();
});


// =========================================
// GENERATE A4 PAGES
// =========================================

function generateCards() {

    // Clear previous cards
    cardsContainer.innerHTML = "";

    let currentPage = null;

    employeeData.forEach(function (employee, index) {

        // New A4 page every 10 cards
        if (index % 10 === 0) {

            currentPage = document.createElement("div");

            currentPage.className = "a4-page";

            currentPage.dataset.page =
                Math.floor(index / 10) + 1;

            cardsContainer.appendChild(currentPage);
        }

        // Clone card
        const card = cardTemplate.cloneNode(true);

        card.classList.remove("card-template");
        card.classList.add("generated-card");

        card.style.display = "block";


        // =================================
        // EMPLOYEE DATA
        // =================================

        const empId = getValue(employee, [
            "EMP ID",
            "EMP_ID",
            "Employee ID",
            "Employee_ID",
            "ID"
        ]);

        const empName = getValue(employee, [
            "EMP NAME",
            "EMP_NAME",
            "Employee Name",
            "Employee_Name",
            "NAME"
        ]);

        const dept = getValue(employee, [
            "DEPT",
            "DEPARTMENT",
            "Department",
            "department"
        ]);

        const subDept = getValue(employee, [
            "SUB-DEPT",
            "SUB DEPT",
            "SUB_DEPT",
            "SUB DEPARTMENT",
            "Sub Department",
            "sub department"
        ]);

        const doj = getValue(employee, [
            "DOJ",
            "Date of Joining",
            "Joining Date"
        ]);

        const gender = getValue(employee, [
            "GENDER",
            "Gender"
        ]);

        const dob = getValue(employee, [
            "DOB",
            "Date of Birth"
        ]);

        const residence = getValue(employee, [
            "RESIDENCE",
            "Residence"
        ]);


        // =================================
        // SET EMPLOYEE DATA
        // =================================

        const empIdElement = card.querySelector(".emp-id");
        const empNameElement = card.querySelector(".emp-name");
        const deptElement = card.querySelector(".dept");
        const subDeptElement = card.querySelector(".sub-dept");
        const dojElement = card.querySelector(".doj");
        const genderElement = card.querySelector(".gender");
        const dobElement = card.querySelector(".dob");

        if (empIdElement) {
            empIdElement.textContent = empId;
        }

        if (empNameElement) {
            empNameElement.textContent = empName;
        }

        if (deptElement) {
            deptElement.textContent = dept;
        }

        if (subDeptElement) {
            subDeptElement.textContent = subDept;
        }

        if (dojElement) {
            dojElement.textContent = formatDate(doj);
        }

        if (genderElement) {
            genderElement.textContent = gender;
        }

        if (dobElement) {
            dobElement.textContent = formatDate(dob);
        }


        // =================================
        // COMPANY DETAILS
        // =================================

        const cardCompany = card.querySelector(".company-name");
        const cardAddress1 = card.querySelector(".address1");
        const cardAddress2 = card.querySelector(".address2");
        const cardAddress3 = card.querySelector(".address3");

        if (cardCompany) {
            cardCompany.textContent = companyName.value;
        }

        if (cardAddress1) {
            cardAddress1.textContent = address1.value;
        }

        if (cardAddress2) {
            cardAddress2.textContent = address2.value;
        }

        if (cardAddress3) {
            cardAddress3.textContent = address3.value;
        }


        // =================================
        // HEADER COLOR
        // =================================

        const header = card.querySelector(".card-header");

        if (header) {

            let residenceColor = "";

            const residenceValue = String(residence);

            if (residenceValue === "DAYSCHOLAR") {
                residenceColor = "#adc086";

            } else if (residenceValue === "LABOUR SHED-II") {
                residenceColor = "#c3add8";

            } else if (residenceValue === "HOSTELLER") {
                residenceColor = "#a2cbd6";

            } else if (residenceValue === "LABOUR SHED-I") {
                residenceColor = "#f3f33d";
            }

            header.style.backgroundColor = residenceColor;
        }


        // =================================
        // EMPLOYEE PHOTO
        // =================================

        const employeePhoto = card.querySelector(".employee-photo");

        const photoPlaceholder = card.querySelector(".photo-placeholder");

        const photoKey = String(empId)
            .trim()
            .toLowerCase();

        if (photos[photoKey]) {

            employeePhoto.src = photos[photoKey];

            employeePhoto.classList.add("loaded");

            if (photoPlaceholder) {
                photoPlaceholder.style.display = "none";
            }

        } else {

            employeePhoto.style.display = "none";

            if (photoPlaceholder) {
                photoPlaceholder.style.display = "flex";
            }
        }


        // =================================
        // ADD CARD TO A4 PAGE
        // =================================

        currentPage.appendChild(card);

    });

    console.log(employeeData.length + " ID cards generated.");
}

// =========================================
// WAIT FOR IMAGES BEFORE PDF EXPORT
// =========================================

async function waitForExportImages(container) {

    const images = Array.from(
        container.querySelectorAll("img")
    );

    await Promise.all(images.map(async (img) => {

        if (!img.src) return;

        if (!img.complete) {

            await new Promise((resolve) => {

                img.addEventListener("load", resolve, {
                    once: true
                });

                img.addEventListener("error", resolve, {
                    once: true
                });
            });
        }

        if (typeof img.decode === "function") {

            try {
                await img.decode();

            } catch (error) {
                // Continue PDF creation if an optional image cannot be decoded.
            }
        }
    }));
}


// =========================================
// PDF GENERATION
// =========================================

pdfBtn.addEventListener("click", async function () {

    const pages = document.querySelectorAll(".a4-page");

    if (pages.length === 0) {
        alert("Please generate ID cards first.");
        return;
    }

    const { jsPDF } = window.jspdf;

    const pdf = new jsPDF({
        orientation: "portrait",
        unit: "mm",
        format: "a4"
    });

    for (let i = 0; i < pages.length; i++) {

        if (i > 0) {
            pdf.addPage();
        }

        // Wait until every logo/photo/signature is fully loaded
        await waitForExportImages(pages[i]);

        const canvas = await html2canvas(pages[i], {

            // Higher render density for sharper text and photos.
            // Card dimensions/alignment are unchanged.
            scale: 6,

            useCORS: true,
            backgroundColor: "#ffffff",
            imageTimeout: 0,
            logging: false
        });

        // PNG keeps the rendered page lossless.
        const image = canvas.toDataURL("image/png");

        pdf.addImage(
            image,
            "PNG",
            0,
            0,
            210,
            297,
            undefined,
            "NONE"
        );
    }

    pdf.save("Employee-ID-Cards-A4.pdf");
});


// =========================================
// CLEAR
// =========================================

clearBtn.addEventListener("click", function () {

    employeeData = [];
    photos = {};

    excelFile.value = "";
    photoFiles.value = "";

    excelStatus.textContent = "No Excel selected";
    photoStatus.textContent = "No photos selected";

    cardsContainer.innerHTML = "";

    cardsContainer.classList.remove("active");

    idCardPage.style.display = "grid";
    idCardPage2.style.display = "none";


    generateBtn.style = "display:";
    btnContainer.style = "bottom: ;display: flex";
    pdfBtn.style.display = "none";
    navBar.style.display = "grid";

});