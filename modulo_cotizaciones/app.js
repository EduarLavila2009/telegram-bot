// ==========================================================================
// LOGICA DE LA APLICACIÓN INTERACTIVA - SUFEVICA COTIZACIONES
// ==========================================================================

document.addEventListener("DOMContentLoaded", () => {
    // 1. Estado Global del Documento
    const state = {
        docType: "cotizacion", // "cotizacion" o "nota"
        currency: "usd",      // "usd" o "ves"
        exchangeRate: 550.00,
        docNumber: "001001",
        docDate: new Date().toISOString().split("T")[0],
        client: {
            name: "",
            address: "",
            rif: "",
            phone: "",
            salesman: "FREDDY LOPEZ",
            saleType: "Contado",
            note: ""
        },
        items: [] // Elementos de la tabla: { code, desc, qty, priceUsd }
    };

    // 2. Elementos del DOM - Entradas de Control
    const elExchangeRate = document.getElementById("exchangeRate");
    const elDocNumber = document.getElementById("docNumber");
    const elDocDate = document.getElementById("docDate");
    const elClientName = document.getElementById("clientName");
    const elClientAddress = document.getElementById("clientAddress");
    const elClientRif = document.getElementById("clientRif");
    const elClientPhone = document.getElementById("clientPhone");
    const elSalesman = document.getElementById("salesman");
    const elSaleType = document.getElementById("saleType");
    const elClientNote = document.getElementById("clientNote");

    // Elementos del Formulario de Ítems
    const elItemCode = document.getElementById("itemCode");
    const elItemDesc = document.getElementById("itemDesc");
    const elItemQty = document.getElementById("itemQty");
    const elItemPrice = document.getElementById("itemPrice");
    const btnAddItem = document.getElementById("addItemBtn");

    // Elementos del DOM - Vista Previa (Hojas)
    const viewDocTypeTitle = document.getElementById("viewDocTypeTitle");
    const viewDocNumber = document.getElementById("viewDocNumber");
    const viewDocDate = document.getElementById("viewDocDate");
    const viewClientName = document.getElementById("viewClientName");
    const viewClientAddress = document.getElementById("viewClientAddress");
    const viewClientRif = document.getElementById("viewClientRif");
    const viewClientPhone = document.getElementById("viewClientPhone");
    const viewSaleType = document.getElementById("viewSaleType");
    const viewSalesman = document.getElementById("viewSalesman");

    const itemsTableBody = document.getElementById("itemsTableBody");
    const colPriceHeader = document.getElementById("colPriceHeader");
    const colTotalHeader = document.getElementById("colTotalHeader");

    const labelSubtotal = document.getElementById("labelSubtotal");
    const labelTotal = document.getElementById("labelTotal");
    const viewSubtotal = document.getElementById("viewSubtotal");
    const viewIva = document.getElementById("viewIva");
    const viewTotal = document.getElementById("viewTotal");

    // Botones de Acción
    const btnClear = document.getElementById("clearBtn");
    const btnDraft = document.getElementById("draftBtn");
    const btnPrint = document.getElementById("printBtn");

    // ==========================================================================
    // FUNCIONES DE FORMATEO Y CONVERSIÓN
    // ==========================================================================

    function formatNumber(num) {
        return Number(num).toLocaleString("es-VE", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        });
    }

    function formatDate(dateStr) {
        if (!dateStr) return "—";
        const parts = dateStr.split("-");
        if (parts.length !== 3) return dateStr;
        return `${parts[2]}/${parts[1]}/${parts[0]}`; // DD/MM/AAAA
    }

    // ==========================================================================
    // RENDERIZADO Y CÁLCULOS
    // ==========================================================================

    function renderDocument() {
        const rate = state.currency === "ves" ? state.exchangeRate : 1;
        const curSymbol = state.currency === "usd" ? "$" : "Bs.";

        // 1. Sincronizar Textos Básicos
        viewDocTypeTitle.textContent = state.docType === "cotizacion" ? "COTIZACIÓN" : "NOTA DE ENTREGA";
        viewDocNumber.textContent = `Nº ${state.docNumber || "S/N"}`;
        viewDocDate.textContent = formatDate(state.docDate);
        viewClientName.textContent = state.client.name || "—";
        viewClientAddress.textContent = state.client.address || "—";
        viewClientRif.textContent = state.client.rif || "—";
        viewClientPhone.textContent = state.client.phone || "—";
        viewSaleType.textContent = state.client.saleType;
        viewSalesman.textContent = state.client.salesman || "—";

        // Sincronizar Nota
        const viewClientNote = document.getElementById("viewClientNote");
        const viewCustomNote = document.getElementById("viewCustomNote");
        if (viewClientNote) {
            viewClientNote.innerHTML = state.docType === "cotizacion"
                ? "<strong>NOTA:</strong> Precios sujetos a cambio sin previo aviso. Esta cotización representa un presupuesto informativo. ¡Gracias por su preferencia!"
                : "<strong>NOTA:</strong> La mercancía viaja por cuenta y riesgo del cliente. Documento de despacho informativo. ¡Gracias por su preferencia!";
        }
        if (viewCustomNote) {
            if (state.client.note && state.client.note.trim() !== "") {
                viewCustomNote.textContent = state.client.note;
                viewCustomNote.style.display = "block";
            } else {
                viewCustomNote.style.display = "none";
            }
        }

        // 2. Encabezados de tabla de precios
        colPriceHeader.textContent = `Precio (${curSymbol})`;
        colTotalHeader.textContent = `Total (${curSymbol})`;
        labelSubtotal.textContent = `Sub-Total (${curSymbol})`;
        labelTotal.textContent = `TOTAL ${state.currency === "usd" ? "$" : "Bs"}`;

        // 3. Renderizar Tabla de Ítems
        itemsTableBody.innerHTML = "";
        
        if (state.items.length === 0) {
            const tr = document.createElement("tr");
            tr.className = "empty-placeholder";
            tr.innerHTML = `<td colspan="6" style="text-align: center; color: #999; padding: 20px;">No hay productos agregados en el documento.</td>`;
            itemsTableBody.appendChild(tr);

            viewSubtotal.textContent = "0,00";
            viewIva.textContent = "0,00";
            viewTotal.textContent = "0,00";
            return;
        }

        let subtotalUsd = 0;

        state.items.forEach((item, index) => {
            const itemPriceConv = item.priceUsd * rate;
            const itemTotalConv = item.qty * itemPriceConv;
            subtotalUsd += (item.qty * item.priceUsd);

            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td class="col-code">${item.code || "—"}</td>
                <td class="col-desc">${item.desc}</td>
                <td class="col-qty">${item.qty}</td>
                <td class="col-price">${formatNumber(itemPriceConv)}</td>
                <td class="col-total">${formatNumber(itemTotalConv)}</td>
                <td class="col-actions no-print">
                    <button class="btn-delete-item" data-index="${index}">🗑️</button>
                </td>
            `;
            itemsTableBody.appendChild(tr);
        });

        // Eventos para botones eliminar
        document.querySelectorAll(".btn-delete-item").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const idx = parseInt(e.target.getAttribute("data-index"));
                deleteItem(idx);
            });
        });

        // 4. Calcular Totales final
        const totalConv = subtotalUsd * rate;

        viewSubtotal.textContent = formatNumber(totalConv);
        viewIva.textContent = "0,00";
        viewTotal.textContent = formatNumber(totalConv);
    }

    // ==========================================================================
    // OPERACIONES
    // ==========================================================================

    function deleteItem(idx) {
        state.items.splice(idx, 1);
        renderDocument();
        saveDraftAuto();
    }

    function addItem() {
        const code = elItemCode.value.trim().toUpperCase();
        const desc = elItemDesc.value.trim();
        const qty = parseFloat(elItemQty.value);
        const price = parseFloat(elItemPrice.value);

        if (!desc) {
            alert("⚠️ Por favor ingresa una descripción para el producto.");
            elItemDesc.focus();
            return;
        }

        if (isNaN(qty) || qty <= 0) {
            alert("⚠️ La cantidad debe ser un número mayor a cero.");
            elItemQty.focus();
            return;
        }

        if (isNaN(price) || price < 0) {
            alert("⚠️ El precio unitario debe ser igual o mayor a cero.");
            elItemPrice.focus();
            return;
        }

        // Agregar al estado
        state.items.push({
            code: code,
            desc: desc,
            qty: qty,
            priceUsd: price
        });

        // Limpiar inputs de ítem
        elItemCode.value = "";
        elItemDesc.value = "";
        elItemQty.value = "1";
        elItemPrice.value = "0.00";

        renderDocument();
        saveDraftAuto();
        elItemCode.focus();
    }

    function clearDocument() {
        if (!confirm("⚠️ ¿Estás seguro de que deseas limpiar todo el documento? Perderás los datos ingresados.")) return;
        state.items = [];
        state.client.name = "";
        state.client.address = "";
        state.client.rif = "";
        state.client.phone = "";
        state.client.note = "";
        state.docNumber = "001001";
        
        elClientName.value = "";
        elClientAddress.value = "";
        elClientRif.value = "";
        elClientPhone.value = "";
        elClientNote.value = "";
        elDocNumber.value = "001001";
        
        renderDocument();
        localStorage.removeItem("sufevica_draft");
    }

    // Guardado en LocalStorage
    function saveDraft() {
        localStorage.setItem("sufevica_draft", JSON.stringify(state));
        alert("💾 ¡Borrador guardado con éxito! Se cargará automáticamente al volver a abrir esta página.");
    }

    function saveDraftAuto() {
        localStorage.setItem("sufevica_draft", JSON.stringify(state));
    }

    function loadDraft() {
        if (state.items && state.items.length > 0) return;
        const raw = localStorage.getItem("sufevica_draft");
        if (!raw) return;
        try {
            const draft = JSON.parse(raw);
            
            state.docType = draft.docType || "cotizacion";
            state.currency = draft.currency || "usd";
            state.exchangeRate = parseFloat(draft.exchangeRate) || 550.00;
            state.docNumber = draft.docNumber || "001001";
            state.docDate = draft.docDate || new Date().toISOString().split("T")[0];
            state.client = draft.client || { name: "", address: "", rif: "", phone: "", salesman: "FREDDY LOPEZ", saleType: "Contado", note: "" };
            state.items = draft.items || [];

            // Sincronizar inputs
            elExchangeRate.value = state.exchangeRate;
            elDocNumber.value = state.docNumber;
            elDocDate.value = state.docDate;
            elClientName.value = state.client.name;
            elClientAddress.value = state.client.address;
            elClientRif.value = state.client.rif;
            elClientPhone.value = state.client.phone;
            elSalesman.value = state.client.salesman;
            elSaleType.value = state.client.saleType;
            elClientNote.value = state.client.note || "";

            // Sincronizar botones activos
            document.querySelectorAll("#docTypeToggle .toggle-btn").forEach(btn => {
                btn.classList.toggle("active", btn.getAttribute("data-type") === state.docType);
            });

            document.querySelectorAll("#currencyToggle .toggle-btn").forEach(btn => {
                btn.classList.toggle("active", btn.getAttribute("data-currency") === state.currency);
            });

            renderDocument();
        } catch (e) {
            console.error("Error al cargar borrador", e);
        }
    }

    // ==========================================================================
    // ASOCIACIÓN DE EVENTOS (DATA BINDING)
    // ==========================================================================

    // Tipo de Documento
    document.querySelectorAll("#docTypeToggle .toggle-btn").forEach(btn => {
        btn.addEventListener("click", (e) => {
            document.querySelectorAll("#docTypeToggle .toggle-btn").forEach(b => b.classList.remove("active"));
            e.target.classList.add("active");
            state.docType = e.target.getAttribute("data-type");
            renderDocument();
            saveDraftAuto();
        });
    });

    // Moneda de Visualización
    document.querySelectorAll("#currencyToggle .toggle-btn").forEach(btn => {
        btn.addEventListener("click", (e) => {
            document.querySelectorAll("#currencyToggle .toggle-btn").forEach(b => b.classList.remove("active"));
            e.target.classList.add("active");
            state.currency = e.target.getAttribute("data-currency");
            renderDocument();
            saveDraftAuto();
        });
    });

    // Cambios en campos cliente/documento
    elDocNumber.addEventListener("input", (e) => {
        state.docNumber = e.target.value;
        renderDocument();
        saveDraftAuto();
    });

    elDocDate.addEventListener("change", (e) => {
        state.docDate = e.target.value;
        renderDocument();
        saveDraftAuto();
    });

    elExchangeRate.addEventListener("input", (e) => {
        state.exchangeRate = parseFloat(e.target.value) || 1.00;
        renderDocument();
        saveDraftAuto();
    });

    elClientName.addEventListener("input", (e) => {
        state.client.name = e.target.value;
        renderDocument();
        saveDraftAuto();
    });

    elClientAddress.addEventListener("input", (e) => {
        state.client.address = e.target.value;
        renderDocument();
        saveDraftAuto();
    });

    elClientRif.addEventListener("input", (e) => {
        state.client.rif = e.target.value;
        renderDocument();
        saveDraftAuto();
    });

    elClientPhone.addEventListener("input", (e) => {
        state.client.phone = e.target.value;
        renderDocument();
        saveDraftAuto();
    });

    elSalesman.addEventListener("input", (e) => {
        state.client.salesman = e.target.value;
        renderDocument();
        saveDraftAuto();
    });

    elClientNote.addEventListener("input", (e) => {
        state.client.note = e.target.value;
        renderDocument();
        saveDraftAuto();
    });

    elSaleType.addEventListener("change", (e) => {
        state.client.saleType = e.target.value;
        renderDocument();
        saveDraftAuto();
    });

    // Botones de Acción
    btnAddItem.addEventListener("click", addItem);
    btnClear.addEventListener("click", clearDocument);
    btnDraft.addEventListener("click", saveDraft);
    
    btnPrint.addEventListener("click", () => {
        window.print();
    });

    // ==========================================================================
    // ARRANQUE INICIAL
    // ==========================================================================
    elDocDate.value = state.docDate;
    loadDraft();

    // Obtener tasa oficial BCV en tiempo real (vía DolarAPI)
    fetch("https://ve.dolarapi.com/v1/dolares/oficial")
        .then(response => response.json())
        .then(data => {
            if (data && data.promedio) {
                const liveRate = parseFloat(data.promedio);
                state.exchangeRate = liveRate;
                elExchangeRate.value = liveRate.toFixed(2);
                renderDocument();
                saveDraftAuto();
            }
        })
        .catch(err => console.error("Error al obtener la tasa oficial del BCV:", err));

    // Predeterminar tipo de documento según parámetro en URL (?type=...) o hash (#...)
    const urlParams = new URLSearchParams(window.location.search);
    let urlType = urlParams.get("type") || urlParams.get("docType");
    if (!urlType && window.location.hash) {
        const hash = window.location.hash.toLowerCase();
        if (hash.includes("nota") || hash.includes("entrega")) {
            urlType = "nota";
        } else if (hash.includes("cotizacion") || hash.includes("presupuesto")) {
            urlType = "cotizacion";
        }
    }
    if (urlType === "nota" || urlType === "cotizacion") {
        state.docType = urlType;
        document.querySelectorAll("#docTypeToggle .toggle-btn").forEach(btn => {
            btn.classList.toggle("active", btn.getAttribute("data-type") === state.docType);
        });
    }

    renderDocument();
});
