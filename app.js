/* ============================================================
   R.C PAY
   Annotation + KML Payroll System
   ============================================================ */

const RC_PAY = {

    profile: {
        name: "",
        id: ""
    },

    google: {
        scriptUrl: "",
        sheetId: "",
        token: ""
    },

    configRows: [],
    workers: [],

    annotation: {
        rows: [],
        result: [],
        total: 0
    },

    kml: {
        rawFeatures: [],
        blocks: [],
        result: [],
        total: 0
    },

    final: {
        rows: [],
        total: 0
    }

};


/* ============================================================
   HELPERS
   ============================================================ */

function $(id) {
    return document.getElementById(id);
}

function money(value) {
    return Number(value || 0).toFixed(2);
}

function number(value) {
    const n = parseFloat(value);
    return Number.isFinite(n) ? n : 0;
}

function normalize(value) {
    return String(value || "")
        .trim()
        .toLowerCase()
        .replace(/[_\-]+/g, " ")
        .replace(/\s+/g, " ");
}

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}


/* ============================================================
   LOCAL STORAGE
   ============================================================ */

function loadSettings() {

    try {

        const profile =
            JSON.parse(localStorage.getItem("rc_profile") || "{}");

        const google =
            JSON.parse(localStorage.getItem("rc_google") || "{}");

        RC_PAY.profile = {
            name: profile.name || "",
            id: profile.id || ""
        };

        RC_PAY.google = {
            scriptUrl: google.scriptUrl || "",
            sheetId: google.sheetId || "",
            token: google.token || ""
        };

    } catch (error) {

        console.error(error);

    }

}

function saveSettings() {

    localStorage.setItem(
        "rc_profile",
        JSON.stringify(RC_PAY.profile)
    );

    localStorage.setItem(
        "rc_google",
        JSON.stringify(RC_PAY.google)
    );

}

function updateProfileUI() {

    const name = RC_PAY.profile.name || "Worker";

    $("profileName").value = RC_PAY.profile.name || "";
    $("profileId").value = RC_PAY.profile.id || "";

    $("sidebarProfileName").textContent = name;
    $("topProfile").textContent = name;
    $("dashWorker").textContent = name;

}

function updateGoogleUI() {

    $("googleScriptUrl").value =
        RC_PAY.google.scriptUrl || "";

    $("googleSheetId").value =
        RC_PAY.google.sheetId || "";

    $("googleToken").value =
        RC_PAY.google.token || "";

}


/* ============================================================
   NAVIGATION
   ============================================================ */

function showPage(pageName) {

    document.querySelectorAll(".page").forEach(page => {
        page.classList.remove("active");
    });

    const page = $("page-" + pageName);

    if (page) {
        page.classList.add("active");
    }

    document.querySelectorAll(".nav-btn").forEach(button => {

        button.classList.toggle(
            "active",
            button.dataset.page === pageName
        );

    });

    const titles = {
        dashboard: "Dashboard",
        annotation: "Annotation Pay",
        kml: "KML Pay",
        final: "Final Pay",
        profile: "Worker Profile",
        google: "Google Sheets"
    };

    $("pageTitle").textContent =
        titles[pageName] || "R.C PAY";

}

document.querySelectorAll(".nav-btn").forEach(button => {

    button.addEventListener("click", () => {

        showPage(button.dataset.page);

    });

});


/* ============================================================
   PROFILE
   ============================================================ */

$("saveProfile").addEventListener("click", () => {

    RC_PAY.profile.name =
        $("profileName").value.trim();

    RC_PAY.profile.id =
        $("profileId").value.trim();

    saveSettings();

    updateProfileUI();

    alert("Profile saved.");

});


/* ============================================================
   GOOGLE SETTINGS
   ============================================================ */

$("saveGoogle").addEventListener("click", () => {

    RC_PAY.google.scriptUrl =
        $("googleScriptUrl").value.trim();

    RC_PAY.google.sheetId =
        $("googleSheetId").value.trim();

    RC_PAY.google.token =
        $("googleToken").value.trim();

    saveSettings();

    alert("Google Sheets settings saved.");

});


$("testGoogle").addEventListener("click", async () => {

    if (!RC_PAY.google.scriptUrl) {

        alert("Enter the Google Apps Script Web App URL first.");
        return;

    }

    try {

        await sendToGoogleSheets({
            action: "test",
            worker: RC_PAY.profile.name || "Test"
        });

        alert(
            "Test request sent. Check your Google Sheet."
        );

    } catch (error) {

        console.error(error);

        alert(
            "Google Sheets test failed: " + error.message
        );

    }

});


/* ============================================================
   CONFIG FILE
   ============================================================ */

$("configFile").addEventListener("change", async event => {

    const file = event.target.files[0];

    if (!file) {
        return;
    }

    try {

        RC_PAY.configRows =
            await readSpreadsheet(file);

        if (!RC_PAY.configRows.length) {

            alert("The configuration sheet is empty.");
            return;

        }

        processConfigSheet();

        alert(
            `Loaded ${RC_PAY.configRows.length} configuration rows.`
        );

    } catch (error) {

        console.error(error);

        alert(
            "Could not read configuration file: " +
            error.message
        );

    }

});


async function readSpreadsheet(file) {

    const extension =
        file.name.split(".").pop().toLowerCase();

    if (extension === "csv") {

        const text =
            await file.text();

        const workbook =
            XLSX.read(text, {
                type: "string"
            });

        const sheet =
            workbook.Sheets[workbook.SheetNames[0]];

        return XLSX.utils.sheet_to_json(
            sheet,
            {
                defval: ""
            }
        );

    }

    const buffer =
        await file.arrayBuffer();

    const workbook =
        XLSX.read(buffer, {
            type: "array"
        });

    const sheet =
        workbook.Sheets[workbook.SheetNames[0]];

    return XLSX.utils.sheet_to_json(
        sheet,
        {
            defval: ""
        }
    );

}


/* ============================================================
   CONFIG PROCESSING
   ============================================================ */

function findColumn(row, possibleNames) {

    const keys =
        Object.keys(row);

    for (const name of possibleNames) {

        const wanted =
            normalize(name);

        const key =
            keys.find(k => normalize(k) === wanted);

        if (key) {
            return key;
        }

    }

    return null;
}


function processConfigSheet() {

    const rows = RC_PAY.configRows;

    const workerColumn =
        findColumn(rows[0], [
            "worker",
            "worker name",
            "name",
            "annotator",
            "user",
            "employee"
        ]);

    const roleColumn =
        findColumn(rows[0], [
            "role",
            "worker type",
            "type"
        ]);

    const assignedColumn =
        findColumn(rows[0], [
            "assigned",
            "assignment",
            "assigned to",
            "worker assigned"
        ]);

    const flagColumn =
        findColumn(rows[0], [
            "flag",
            "annotation",
            "annotation type",
            "label",
            "class",
            "feature"
        ]);

    const rateColumn =
        findColumn(rows[0], [
            "rate",
            "pay",
            "price",
            "unit rate"
        ]);

    const uniqueWorkers =
        new Set();

    rows.forEach(row => {

        let worker = "";

        if (workerColumn) {
            worker = String(row[workerColumn]).trim();
        }

        if (worker) {
            uniqueWorkers.add(worker);
        }

    });

    RC_PAY.workers =
        Array.from(uniqueWorkers).sort();

    populateWorkerSelect(
        "annotationWorker",
        RC_PAY.workers
    );

    populateWorkerSelect(
        "kmlWorker",
        RC_PAY.workers
    );

    renderAnnotationRates(
        rows,
        flagColumn,
        rateColumn
    );

}


/* ============================================================
   WORKER SELECT
   ============================================================ */

function populateWorkerSelect(id, workers) {

    const select = $(id);

    select.innerHTML =
        `<option value="">Select worker</option>`;

    workers.forEach(worker => {

        const option =
            document.createElement("option");

        option.value = worker;
        option.textContent = worker;

        select.appendChild(option);

    });

}


/* ============================================================
   ANNOTATION RATES
   ============================================================ */

const ANNOTATION_TYPES = [
    "Location Description",
    "Bus Stop",
    "Bus Lane",
    "Bike Lane",
    "R0",
    "L0",
    "Other"
];


function renderAnnotationRates(
    rows,
    flagColumn,
    rateColumn
) {

    const container =
        $("annotationRates");

    container.innerHTML = "";

    const detectedRates = {};

    rows.forEach(row => {

        if (!flagColumn || !rateColumn) {
            return;
        }

        const flag =
            classifyAnnotation(row[flagColumn]);

        const rate =
            number(row[rateColumn]);

        if (
            flag &&
            !detectedRates[flag] &&
            rate > 0
        ) {

            detectedRates[flag] = rate;

        }

    });

    ANNOTATION_TYPES.forEach(type => {

        const id =
            "annotationRate_" +
            type.toLowerCase()
                .replaceAll(" ", "_");

        const wrapper =
            document.createElement("div");

        wrapper.className = "field";

        wrapper.innerHTML = `
            <label>${escapeHtml(type)} Rate</label>
            <input
                id="${id}"
                type="number"
                step="0.01"
                value="${detectedRates[type] || 0}"
            >
        `;

        container.appendChild(wrapper);

    });

}


function getAnnotationRates() {

    const rates = {};

    ANNOTATION_TYPES.forEach(type => {

        const id =
            "annotationRate_" +
            type.toLowerCase()
                .replaceAll(" ", "_");

        rates[type] =
            number($(id)?.value);

    });

    return rates;

}


/* ============================================================
   CLASSIFY ANNOTATION
   ============================================================ */

function classifyAnnotation(value) {

    const text =
        normalize(value);

    if (!text) {
        return "Other";
    }

    if (
        text.includes("bus stop") ||
        text.includes("busstop") ||
        text === "stop"
    ) {
        return "Bus Stop";
    }

    if (
        text.includes("bus lane") ||
        text.includes("buslane")
    ) {
        return "Bus Lane";
    }

    if (
        text.includes("bike lane") ||
        text.includes("bikelane") ||
        text.includes("cycle lane")
    ) {
        return "Bike Lane";
    }

    if (
        text === "r0" ||
        text.includes("r0 ")
    ) {
        return "R0";
    }

    if (
        text === "l0" ||
        text.includes("l0 ")
    ) {
        return "L0";
    }

    if (
        text.includes("location") ||
        text.includes("description") ||
        text.includes("location description")
    ) {
        return "Location Description";
    }

    return "Other";

}


/* ============================================================
   ASSIGNMENT CHECK
   ============================================================ */

function isAssigned(row, selectedWorker) {

    if (!selectedWorker) {
        return false;
    }

    const keys =
        Object.keys(row);

    const workerKeys =
        keys.filter(key => {

            const n =
                normalize(key);

            return (
                n.includes("worker") ||
                n === "name" ||
                n.includes("annotator") ||
                n.includes("assigned")
            );

        });

    if (!workerKeys.length) {

        return true;

    }

    return workerKeys.some(key => {

        const value =
            String(row[key] || "").trim();

        return (
            normalize(value) ===
            normalize(selectedWorker)
        );

    });

}


/* ============================================================
   ANNOTATION CALCULATION
   ============================================================ */

$("calculateAnnotation")
    .addEventListener("click", calculateAnnotationPay);


function calculateAnnotationPay() {

    if (!RC_PAY.configRows.length) {

        alert(
            "Upload the configuration sheet first."
        );

        return;

    }

    const worker =
        $("annotationWorker").value;

    if (!worker) {

        alert("Select a worker.");
        return;

    }

    const rates =
        getAnnotationRates();

    const firstRow =
        RC_PAY.configRows[0];

    const flagColumn =
        findColumn(firstRow, [
            "flag",
            "annotation",
            "annotation type",
            "label",
            "class",
            "feature"
        ]);

    if (!flagColumn) {

        alert(
            "Could not find an annotation/flag column."
        );

        return;

    }

    const results = {};

    RC_PAY.configRows.forEach(row => {

        if (!isAssigned(row, worker)) {
            return;
        }

        const type =
            classifyAnnotation(row[flagColumn]);

        if (!results[type]) {

            results[type] = {
                quantity: 0,
                rate: rates[type] || 0,
                assigned: true,
                pay: 0
            };

        }

        results[type].quantity++;

    });

    Object.keys(results).forEach(type => {

        results[type].rate =
            rates[type] || 0;

        results[type].pay =
            results[type].quantity *
            results[type].rate;

    });

    const output =
        Object.entries(results).map(
            ([type, data]) => ({
                Worker: worker,
                Type: "Annotation",
                Annotation: type,
                Quantity: data.quantity,
                Rate: data.rate,
                Assigned: data.assigned ? "YES" : "NO",
                Pay: data.pay
            })
        );

    RC_PAY.annotation.result =
        output;

    RC_PAY.annotation.total =
        output.reduce(
            (sum, row) =>
                sum + number(row.Pay),
            0
        );

    renderAnnotationResult();

    updateDashboard();

}


function renderAnnotationResult() {

    const container =
        $("annotationResult");

    if (!RC_PAY.annotation.result.length) {

        container.innerHTML = `
            <div class="alert alert-warning">
                No annotation records were found for this worker.
            </div>
        `;

        return;

    }

    let html = `

        <div class="section">

            <h3>Annotation Pay Result</h3>

            <div class="table-wrapper">

                <table>

                    <thead>

                        <tr>
                            <th>Worker</th>
                            <th>Type</th>
                            <th>Annotation</th>
                            <th>Quantity</th>
                            <th>Rate</th>
                            <th>Assigned</th>
                            <th>Pay</th>
                        </tr>

                    </thead>

                    <tbody>
    `;

    RC_PAY.annotation.result.forEach(row => {

        html += `

            <tr>

                <td>${escapeHtml(row.Worker)}</td>

                <td>${escapeHtml(row.Type)}</td>

                <td>${escapeHtml(row.Annotation)}</td>

                <td>${row.Quantity}</td>

                <td>${money(row.Rate)}</td>

                <td>
                    <span class="badge badge-success">
                        YES
                    </span>
                </td>

                <td class="money">
                    ${money(row.Pay)}
                </td>

            </tr>

        `;

    });

    html += `

                    </tbody>

                </table>

            </div>

            <br>

            <div class="pay-total">

                <div>
                    <div class="pay-total-label">
                        Annotation Pay
                    </div>

                    <div class="pay-total-value">
                        ${money(RC_PAY.annotation.total)}
                    </div>
                </div>

            </div>

        </div>

    `;

    container.innerHTML = html;

}


/* ============================================================
   KML FILE
   ============================================================ */

$("kmlFile").addEventListener(
    "change",
    handleKmlFile
);


async function handleKmlFile(event) {

    const file =
        event.target.files[0];

    if (!file) {
        return;
    }

    try {

        let xmlText = "";

        if (
            file.name
                .toLowerCase()
                .endsWith(".kmz")
        ) {

            const zip =
                await JSZip.loadAsync(file);

            const kmlEntry =
                Object.keys(zip.files)
                    .find(name =>
                        name.toLowerCase().endsWith(".kml")
                    );

            if (!kmlEntry) {

                throw new Error(
                    "No KML file found inside KMZ."
                );

            }

            xmlText =
                await zip.files[kmlEntry]
                    .async("text");

        } else {

            xmlText =
                await file.text();

        }

        RC_PAY.kml.rawFeatures =
            parseKml(xmlText);

        RC_PAY.kml.blocks =
            groupKmlIntoBlocks(
                RC_PAY.kml.rawFeatures
            );

        renderKmlPreview();

        $("dashBlocks").textContent =
            RC_PAY.kml.blocks.length;

        alert(
            `Loaded ${RC_PAY.kml.rawFeatures.length} KML features in ${RC_PAY.kml.blocks.length} blocks.`
        );

    } catch (error) {

        console.error(error);

        alert(
            "KML processing failed: " +
            error.message
        );

    }

}


/* ============================================================
   KML PARSER
   ============================================================ */

function getXmlText(parent, tagName) {

    const element =
        parent.getElementsByTagName(tagName)[0];

    return element
        ? element.textContent.trim()
        : "";

}


function parseKml(xmlText) {

    const parser =
        new DOMParser();

    const xml =
        parser.parseFromString(
            xmlText,
            "application/xml"
        );

    const parseError =
        xml.querySelector("parsererror");

    if (parseError) {

        throw new Error(
            "Invalid KML XML."
        );

    }

    const placemarks =
        Array.from(
            xml.getElementsByTagName("Placemark")
        );

    return placemarks.map(
        (placemark, index) => {

            const name =
                getXmlText(
                    placemark,
                    "name"
                );

            const description =
                getXmlText(
                    placemark,
                    "description"
                );

            const folder =
                findParentFolder(
                    placemark
                );

            const coordinates =
                getXmlText(
                    placemark,
                    "coordinates"
                );

            const extendedData =
                readExtendedData(
                    placemark
                );

            const rawText =
                [
                    name,
                    description,
                    folder,
                    JSON.stringify(extendedData)
                ]
                .join(" ");

            const type =
                classifyKmlFeature(
                    rawText
                );

            const block =
                detectBlock(
                    name,
                    description,
                    folder,
                    extendedData
                );

            const location =
                detectLocation(
                    name,
                    description,
                    folder,
                    extendedData
                );

            return {

                index,

                name,

                description,

                folder,

                coordinates,

                extendedData,

                type,

                block,

                location,

                rawText

            };

        }
    );

}


function findParentFolder(placemark) {

    let parent =
        placemark.parentElement;

    while (parent) {

        if (
            parent.tagName &&
            parent.tagName.toLowerCase() ===
            "folder"
        ) {

            const name =
                getXmlText(
                    parent,
                    "name"
                );

            if (name) {
                return name;
            }

        }

        parent =
            parent.parentElement;

    }

    return "";

}


function readExtendedData(placemark) {

    const result = {};

    const data =
        placemark.getElementsByTagName(
            "ExtendedData"
        )[0];

    if (!data) {
        return result;
    }

    const datas =
        Array.from(
            data.getElementsByTagName("Data")
        );

    datas.forEach(item => {

        const key =
            item.getAttribute("name");

        const value =
            getXmlText(
                item,
                "value"
            );

        if (key) {
            result[key] = value;
        }

    });

    const simpleData =
        Array.from(
            data.getElementsByTagName(
                "SimpleData"
            )
        );

    simpleData.forEach(item => {

        const key =
            item.getAttribute("name");

        const value =
            item.textContent.trim();

        if (key) {
            result[key] = value;
        }

    });

    return result;

}


/* ============================================================
   KML CLASSIFICATION
   ============================================================ */

function classifyKmlFeature(text) {

    const n =
        normalize(text);

    if (
        n.includes("bus stop") ||
        n.includes("busstop")
    ) {

        return "Bus Stop";

    }

    if (
        n.includes("bus lane") ||
        n.includes("buslane")
    ) {

        return "Bus Lane";

    }

    if (
        n.includes("bike lane") ||
        n.includes("bikelane") ||
        n.includes("cycle lane")
    ) {

        return "Bike Lane";

    }

    if (
        /\br0\b/i.test(text) ||
        n === "r0"
    ) {

        return "R0";

    }

    if (
        /\bl0\b/i.test(text) ||
        n === "l0"
    ) {

        return "L0";

    }

    if (
        n.includes("location description") ||
        n.includes("location") ||
        n.includes("description")
    ) {

        return "Location Description";

    }

    return "Other";

}


/* ============================================================
   BLOCK DETECTION
   ============================================================ */

function detectBlock(
    name,
    description,
    folder,
    extendedData
) {

    const values = [

        name,

        description,

        folder,

        ...Object.values(
            extendedData || {}
        )

    ];

    const text =
        values
            .filter(Boolean)
            .join(" ");

    /*
       Supported examples:

       Block 001
       Block-001
       Block_001
       block_id=001
       block: 001
       segment 001
       road 001
    */

    const patterns = [

        /\bblock[\s_\-:#]*(\w+)\b/i,

        /\bblock[\s_\-]*id[\s:#=]*(\w+)\b/i,

        /\bsegment[\s_\-:#]*(\w+)\b/i,

        /\broad[\s_\-:#]*(\w+)\b/i

    ];

    for (const pattern of patterns) {

        const match =
            text.match(pattern);

        if (match && match[1]) {

            return "BLOCK-" +
                match[1];

        }

    }

    /*
       If no explicit block exists,
       use the folder name.
    */

    if (folder) {

        return folder.trim();

    }

    /*
       Last fallback:
       each feature becomes its own block.
    */

    return "BLOCK-UNSPECIFIED";

}


/* ============================================================
   LOCATION DESCRIPTION
   ============================================================ */

function detectLocation(
    name,
    description,
    folder,
    extendedData
) {

    const possibleKeys = [
        "location",
        "location description",
        "location_description",
        "description",
        "place",
        "road",
        "street"
    ];

    for (const key of possibleKeys) {

        const found =
            Object.keys(
                extendedData || {}
            ).find(k =>
                normalize(k) === normalize(key)
            );

        if (
            found &&
            extendedData[found]
        ) {

            return extendedData[found];

        }

    }

    if (description) {

        return stripHtml(description);

    }

    if (name) {

        return name;

    }

    return "";

}


function stripHtml(value) {

    const div =
        document.createElement("div");

    div.innerHTML =
        String(value || "");

    return div.textContent.trim();

}


/* ============================================================
   GROUP KML BY BLOCK
   ============================================================ */

function groupKmlIntoBlocks(features) {

    const map =
        new Map();

    features.forEach(feature => {

        const blockName =
            feature.block ||
            "BLOCK-UNSPECIFIED";

        if (!map.has(blockName)) {

            map.set(
                blockName,
                {

                    block: blockName,

                    location: "",

                    busStop: 0,

                    busLane: 0,

                    bikeLane: 0,

                    r0: 0,

                    l0: 0,

                    locationDescription: 0,

                    other: 0,

                    otherNames: [],

                    features: []

                }
            );

        }

        const block =
            map.get(blockName);

        block.features.push(feature);

        if (
            !block.location &&
            feature.location
        ) {

            block.location =
                feature.location;

        }

        switch (feature.type) {

            case "Bus Stop":
                block.busStop++;
                break;

            case "Bus Lane":
                block.busLane++;
                break;

            case "Bike Lane":
                block.bikeLane++;
                break;

            case "R0":
                block.r0++;
                break;

            case "L0":
                block.l0++;
                break;

            case "Location Description":
                block.locationDescription++;
                break;

            default:

                block.other++;

                if (
                    feature.name &&
                    !block.otherNames.includes(
                        feature.name
                    )
                ) {

                    block.otherNames.push(
                        feature.name
                    );

                }

                break;

        }

    });

    return Array.from(map.values());

}


/* ============================================================
   KML PREVIEW
   ============================================================ */

function renderKmlPreview() {

    const existing =
        $("kmlResult");

    if (!RC_PAY.kml.blocks.length) {

        existing.innerHTML = "";
        return;

    }

    let html = `

        <div class="section">

            <h3>KML Block Preview</h3>

            <div class="section-description">
                Each block below will become one row in the final report.
            </div>

            <div class="kml-preview">

    `;

    RC_PAY.kml.blocks
        .slice(0, 100)
        .forEach(block => {

            html += `

                <div class="block-card">

                    <div class="block-title">
                        ${escapeHtml(block.block)}
                    </div>

                    <div style="font-size:12px;color:#6b7280;margin-bottom:8px;">
                        ${escapeHtml(block.location || "No location description")}
                    </div>

                    <div class="annotation-pills">

                        <span class="annotation-pill">
                            Bus Stop: ${block.busStop}
                        </span>

                        <span class="annotation-pill">
                            Bus Lane: ${block.busLane}
                        </span>

                        <span class="annotation-pill">
                            Bike Lane: ${block.bikeLane}
                        </span>

                        <span class="annotation-pill">
                            R0: ${block.r0}
                        </span>

                        <span class="annotation-pill">
                            L0: ${block.l0}
                        </span>

                        <span class="annotation-pill">
                            Location: ${block.locationDescription}
                        </span>

                        <span class="annotation-pill">
                            Additional: ${block.other}
                        </span>

                    </div>

                </div>

            `;

        });

    html += `

            </div>

        </div>

    `;

    existing.innerHTML = html;

}


/* ============================================================
   KML CALCULATION
   ============================================================ */

$("calculateKml")
    .addEventListener(
        "click",
        calculateKmlPay
    );


function calculateKmlPay() {

    if (!RC_PAY.kml.blocks.length) {

        alert(
            "Upload a KML/KMZ file first."
        );

        return;

    }

    const worker =
        $("kmlWorker").value;

    if (!worker) {

        alert(
            "Select the worker whose KML work should be calculated."
        );

        return;

    }

    const rates = {

        "Bus Stop":
            number($("rateBusStop").value),

        "Bus Lane":
            number($("rateBusLane").value),

        "Bike Lane":
            number($("rateBikeLane").value),

        "R0":
            number($("rateR0").value),

        "L0":
            number($("rateL0").value),

        "Location Description":
            number($("rateLocation").value),

        "Other":
            number($("rateOther").value)

    };

    const payUnknown =
        $("payUnknown").checked;

    const result = [];

    RC_PAY.kml.blocks.forEach(block => {

        const busStopPay =
            block.busStop *
            rates["Bus Stop"];

        const busLanePay =
            block.busLane *
            rates["Bus Lane"];

        const bikeLanePay =
            block.bikeLane *
            rates["Bike Lane"];

        const r0Pay =
            block.r0 *
            rates["R0"];

        const l0Pay =
            block.l0 *
            rates["L0"];

        const locationPay =
            block.locationDescription *
            rates["Location Description"];

        const otherPay =
            payUnknown
                ? block.other * rates["Other"]
                : 0;

        const totalPay =
            busStopPay +
            busLanePay +
            bikeLanePay +
            bikeLanePay * 0 +
            r0Pay +
            l0Pay +
            locationPay +
            otherPay;

        result.push({

            Worker: worker,

            Type: "KML",

            Block: block.block,

            Location:
                block.location,

            BusStop:
                block.busStop,

            BusLane:
                block.busLane,

            BikeLane:
                block.bikeLane,

            R0:
                block.r0,

            L0:
                block.l0,

            LocationDescription:
                block.locationDescription,

            Additional:
                block.other,

            AdditionalNames:
                block.otherNames.join("; "),

            BusStopPay:
                busStopPay,

            BusLanePay:
                busLanePay,

            BikeLanePay:
                bikeLanePay,

            R0Pay:
                r0Pay,

            L0Pay:
                l0Pay,

            LocationPay:
                locationPay,

            AdditionalPay:
                otherPay,

            Pay:
                totalPay

        });

    });

    RC_PAY.kml.result =
        result;

    RC_PAY.kml.total =
        result.reduce(
            (sum, row) =>
                sum + number(row.Pay),
            0
        );

    renderKmlResult();

    updateDashboard();

}


function renderKmlResult() {

    const container =
        $("kmlResult");

    let html = `

        <div class="section">

            <h3>KML Pay Report</h3>

            <div class="section-description">
                One row = one block. Bus Stop, Bus Lane, Bike Lane,
                R0 and L0 remain together on that block's row.
            </div>

            <div class="table-wrapper">

                <table>

                    <thead>

                        <tr>

                            <th>Worker</th>
                            <th>Block</th>
                            <th>Location</th>
                            <th>Bus Stop</th>
                            <th>Bus Lane</th>
                            <th>Bike Lane</th>
                            <th>R0</th>
                            <th>L0</th>
                            <th>Location Description</th>
                            <th>Additional</th>
                            <th>Total Pay</th>

                        </tr>

                    </thead>

                    <tbody>

    `;

    RC_PAY.kml.result.forEach(row => {

        html += `

            <tr>

                <td>${escapeHtml(row.Worker)}</td>

                <td>${escapeHtml(row.Block)}</td>

                <td>${escapeHtml(row.Location)}</td>

                <td>${row.BusStop}</td>

                <td>${row.BusLane}</td>

                <td>${row.BikeLane}</td>

                <td>${row.R0}</td>

                <td>${row.L0}</td>

                <td>${row.LocationDescription}</td>

                <td>

                    ${
                        row.Additional > 0
                        ? `<span class="badge badge-warning">
                            ${row.Additional}
                           </span>`
                        : "0"
                    }

                </td>

                <td class="money">
                    ${money(row.Pay)}
                </td>

            </tr>

        `;

    });

    html += `

                    </tbody>

                </table>

            </div>

            <br>

            <div class="pay-total">

                <div>

                    <div class="pay-total-label">
                        KML Pay
                    </div>

                    <div class="pay-total-value">
                        ${money(RC_PAY.kml.total)}
                    </div>

                </div>

                <div>

                    <div class="pay-total-label">
                        Blocks
                    </div>

                    <div class="pay-total-value">
                        ${RC_PAY.kml.result.length}
                    </div>

                </div>

            </div>

        </div>

    `;

    container.innerHTML = html;

}


/* ============================================================
   FINAL PAY
   ============================================================ */

function buildFinalReport() {

    const rows = [];

    RC_PAY.annotation.result.forEach(row => {

        rows.push({

            Worker:
                row.Worker,

            Source:
                "Annotation",

            Block:
                "",

            Location:
                "",

            Annotation:
                row.Annotation,

            Quantity:
                row.Quantity,

            Rate:
                row.Rate,

            Assigned:
                row.Assigned,

            Pay:
                row.Pay

        });

    });

    RC_PAY.kml.result.forEach(row => {

        rows.push({

            Worker:
                row.Worker,

            Source:
                "KML",

            Block:
                row.Block,

            Location:
                row.Location,

            Annotation:
                "Bus Stop / Bus Lane / Bike Lane / R0 / L0",

            Quantity:
                row.BusStop +
                row.BusLane +
                row.BikeLane +
                row.R0 +
                row.L0 +
                row.LocationDescription +
                row.Additional,

            Rate:
                "",

            Assigned:
                "YES",

            Pay:
                row.Pay

        });

    });

    RC_PAY.final.rows =
        rows;

    RC_PAY.final.total =
        rows.reduce(
            (sum, row) =>
                sum + number(row.Pay),
            0
        );

}


function renderFinalSummary() {

    buildFinalReport();

    const container =
        $("finalSummary");

    const worker =
        RC_PAY.profile.name ||
        $("annotationWorker").value ||
        $("kmlWorker").value ||
        "Worker";

    const annotationTotal =
        RC_PAY.annotation.total;

    const kmlTotal =
        RC_PAY.kml.total;

    let html = `

        <div class="cards">

            <div class="card">

                <div class="stat-label">
                    Worker
                </div>

                <div class="stat-value"
                     style="font-size:20px;">
                    ${escapeHtml(worker)}
                </div>

            </div>

            <div class="card">

                <div class="stat-label">
                    Annotation Pay
                </div>

                <div class="stat-value">
                    ${money(annotationTotal)}
                </div>

            </div>

            <div class="card">

                <div class="stat-label">
                    KML Pay
                </div>

                <div class="stat-value">
                    ${money(kmlTotal)}
                </div>

            </div>

            <div class="card">

                <div class="stat-label">
                    Total Pay
                </div>

                <div class="stat-value">
                    ${money(RC_PAY.final.total)}
                </div>

            </div>

        </div>

        <div class="pay-total">

            <div>

                <div class="pay-total-label">
                    FINAL PAY
                </div>

                <div class="pay-total-value">
                    ${money(RC_PAY.final.total)}
                </div>

            </div>

        </div>

        <br>

        <div class="table-wrapper">

            <table>

                <thead>

                    <tr>

                        <th>Worker</th>
                        <th>Source</th>
                        <th>Block</th>
                        <th>Location</th>
                        <th>Annotation</th>
                        <th>Quantity</th>
                        <th>Rate</th>
                        <th>Assigned</th>
                        <th>Pay</th>

                    </tr>

                </thead>

                <tbody>

    `;

    RC_PAY.final.rows.forEach(row => {

        html += `

            <tr>

                <td>${escapeHtml(row.Worker)}</td>

                <td>${escapeHtml(row.Source)}</td>

                <td>${escapeHtml(row.Block)}</td>

                <td>${escapeHtml(row.Location)}</td>

                <td>${escapeHtml(row.Annotation)}</td>

                <td>${row.Quantity}</td>

                <td>
                    ${row.Rate === ""
                        ? "-"
                        : money(row.Rate)}
                </td>

                <td>
                    <span class="badge badge-success">
                        ${row.Assigned}
                    </span>
                </td>

                <td class="money">
                    ${money(row.Pay)}
                </td>

            </tr>

        `;

    });

    html += `

                </tbody>

            </table>

        </div>

    `;

    container.innerHTML =
        html;

}


/* ============================================================
   FINAL PAGE ACTIVATION
   ============================================================ */

document
    .querySelector('[data-page="final"]')
    .addEventListener(
        "click",
        renderFinalSummary
    );


/* ============================================================
   CSV
   ============================================================ */

$("downloadCsv")
    .addEventListener(
        "click",
        downloadFinalCsv
    );


function csvEscape(value) {

    const stringValue =
        String(value ?? "");

    if (
        stringValue.includes(",") ||
        stringValue.includes('"') ||
        stringValue.includes("\n")
    ) {

        return '"' +
            stringValue.replaceAll('"', '""') +
            '"';

    }

    return stringValue;

}


function downloadFinalCsv() {

    buildFinalReport();

    if (!RC_PAY.final.rows.length) {

        alert(
            "There is no calculated pay report yet."
        );

        return;

    }

    const columns = [

        "Worker",
        "Source",
        "Block",
        "Location",
        "Annotation",
        "Quantity",
        "Rate",
        "Assigned",
        "Pay"

    ];

    const lines = [];

    lines.push(
        columns.join(",")
    );

    RC_PAY.final.rows.forEach(row => {

        lines.push(
            columns
                .map(column =>
                    csvEscape(row[column])
                )
                .join(",")
        );

    });

    lines.push("");

    lines.push(
        [
            "",
            "",
            "",
            "",
            "FINAL PAY",
            "",
            "",
            "",
            money(RC_PAY.final.total)
        ].join(",")
    );

    const blob =
        new Blob(
            [lines.join("\n")],
            {
                type: "text/csv;charset=utf-8;"
            }
        );

    const url =
        URL.createObjectURL(blob);

    const link =
        document.createElement("a");

    link.href = url;

    link.download =
        `RC_PAY_${safeFileName(
            RC_PAY.profile.name || "worker"
        )}.csv`;

    document.body.appendChild(link);

    link.click();

    link.remove();

    URL.revokeObjectURL(url);

}


function safeFileName(value) {

    return String(value)
        .replace(/[^a-z0-9_\-]+/gi, "_")
        .replace(/^_+|_+$/g, "");

}


/* ============================================================
   GOOGLE SHEETS
   ============================================================ */

$("sendGoogle")
    .addEventListener(
        "click",
        sendFinalToGoogle
    );


async function sendFinalToGoogle() {

    buildFinalReport();

    if (!RC_PAY.final.rows.length) {

        alert(
            "Calculate pay before sending the report."
        );

        return;

    }

    if (!RC_PAY.google.scriptUrl) {

        alert(
            "Configure Google Sheets first."
        );

        showPage("google");

        return;

    }

    try {

        await sendToGoogleSheets({

            action: "saveReport",

            sheetId:
                RC_PAY.google.sheetId,

            token:
                RC_PAY.google.token,

            worker: {

                name:
                    RC_PAY.profile.name,

                id:
                    RC_PAY.profile.id

            },

            annotationTotal:
                RC_PAY.annotation.total,

            kmlTotal:
                RC_PAY.kml.total,

            finalTotal:
                RC_PAY.final.total,

            rows:
                RC_PAY.final.rows,

            kmlRows:
                RC_PAY.kml.result

        });

        alert(
            "Report sent to Google Sheets."
        );

    } catch (error) {

        console.error(error);

        alert(
            "Could not send report: " +
            error.message
        );

    }

}


async function sendToGoogleSheets(payload) {

    const url =
        RC_PAY.google.scriptUrl;

    if (!url) {

        throw new Error(
            "Google Apps Script URL is missing."
        );

    }

    const body = {

        ...payload,

        token:
            RC_PAY.google.token,

        sheetId:
            RC_PAY.google.sheetId

    };

    /*
       no-cors is intentional for Apps Script Web Apps.
       The request is sent to Apps Script but the browser
       cannot inspect the response.
    */

    await fetch(
        url,
        {

            method: "POST",

            mode: "no-cors",

            headers: {
                "Content-Type":
                    "text/plain;charset=utf-8"
            },

            body:
                JSON.stringify(body)

        }
    );

}


/* ============================================================
   DASHBOARD
   ============================================================ */

function updateDashboard() {

    $("dashAnnotations").textContent =
        RC_PAY.annotation.result.length;

    $("dashBlocks").textContent =
        RC_PAY.kml.blocks.length;

    $("dashPay").textContent =
        money(
            RC_PAY.annotation.total +
            RC_PAY.kml.total
        );

}


/* ============================================================
   CLEAR
   ============================================================ */

$("clearAnnotation")
    .addEventListener(
        "click",
        () => {

            RC_PAY.configRows = [];
            RC_PAY.workers = [];

            RC_PAY.annotation.result = [];
            RC_PAY.annotation.total = 0;

            $("configFile").value = "";

            $("annotationWorker").innerHTML =
                `<option value="">Upload configuration first</option>`;

            $("kmlWorker").innerHTML =
                `<option value="">Select worker</option>`;

            $("annotationRates").innerHTML = "";

            $("annotationResult").innerHTML = "";

            updateDashboard();

        }
    );


$("clearKml")
    .addEventListener(
        "click",
        () => {

            RC_PAY.kml.rawFeatures = [];
            RC_PAY.kml.blocks = [];
            RC_PAY.kml.result = [];
            RC_PAY.kml.total = 0;

            $("kmlFile").value = "";

            $("kmlResult").innerHTML = "";

            updateDashboard();

        }
    );


/* ============================================================
   INITIALIZE
   ============================================================ */

loadSettings();

updateProfileUI();

updateGoogleUI();

updateDashboard();
