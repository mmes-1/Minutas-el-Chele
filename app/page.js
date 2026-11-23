"use client";
import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  ShoppingCart,
  Package,
  History,
  Wallet,
  Lock,
  Unlock,
  Download,
  Plus,
  Minus,
  Trash2,
  DollarSign,
  CreditCard,
  ArrowRightLeft,
  Search,
  CheckCircle2,
  EyeOff,
  ClipboardCheck,
  Calendar,
  AlertCircle,
  Target,
  Pencil,
  X,
  Layers,
} from "lucide-react";

const OWNER_PASSWORD = "Negocio2026";

const INITIAL_MENU = [
  {
    id: 1,
    name: "Ceviminuta",
    category: "Minutas",
    price: 5.0,
    stock: 400,
    description:
      "Minuta de limon, sal, chile, chamoy, salsa negra, tajin, alguashte, nachos, mango o pepino, chimol, camarones",
    recipe: [],
  },
  {
    id: 2,
    name: "Limon y fruta",
    category: "Minutas",
    price: 3.0,
    stock: 400,
    description:
      "Minuta de limon, sal, chile, chamoy, salsa negra, tajin, alguashte, nachos, mango o pepino",
    recipe: [],
  },
  {
    id: 3,
    name: "Tadicional con gomitas",
    category: "Minutas",
    price: 3.0,
    stock: 400,
    description:
      "Minuta de Jarabes de tu elección, leche condensada, gomitas y pixie",
    recipe: [],
  },
  {
    id: 4,
    name: "Tadicional con fruta",
    category: "Minutas",
    price: 2.0,
    stock: 400,
    description:
      "Minuta de Jarabes de tu elección con fruta picada (guineo y piña) y jalea de tamarindo o leche condensada",
    recipe: [],
  },
  {
    id: 5,
    name: "Tradicional de limon",
    category: "Minutas",
    price: 1.5,
    stock: 400,
    description:
      "Minuta de limon, sal, chile, chamoy, salsa negra, tajin, alguashte",
    recipe: [],
  },
  {
    id: 6,
    name: "Licuado de guineo",
    category: "Licuados",
    price: 2.0,
    stock: 600,
    description: "Licuado de guineo preparado con leche",
    recipe: [],
  },
  {
    id: 7,
    name: "Licuado de fresa",
    category: "Licuados",
    price: 2.0,
    stock: 600,
    description: "Licuado de fresa preparado con leche",
    recipe: [],
  },
  {
    id: 8,
    name: "Licuado de Oreo",
    category: "Licuados",
    price: 2.0,
    stock: 600,
    description: "Licuado de galleta Oreo preparado con leche",
    recipe: [],
  },
];

const INITIAL_SUPPLIES = [
  { id: "s1", name: "vasos de licuados", type: "Empaques", stock: 600 },
  { id: "s2", name: "vasos de minutas", type: "Empaques", stock: 300 },
  { id: "s3", name: "bandeja 11,1x11,1,x6,3", type: "Empaques", stock: 400 },
  { id: "s4", name: "stickers", type: "Empaques", stock: 0 },
  { id: "s5", name: "tenedores", type: "Empaques", stock: 300 },
  { id: "s6", name: "cucharas", type: "Empaques", stock: 300 },
  { id: "s7", name: "pajillas sin emboltorio azul", type: "Empaques", stock: 300 },
  { id: "s8", name: "pajillas con emboltorio", type: "Empaques", stock: 340 },
  { id: "s9", name: "servilletas", type: "Empaques", stock: 1000 },
  {
    id: "s10",
    name: "Bolsas para guardar la fruta picada",
    type: "Empaques",
    stock: 300,
  },
  { id: "s11", name: "Jarabes", type: "Materia Prima", stock: 8 },
  { id: "s12", name: "Jalea de tamarindo", type: "Materia Prima", stock: 0 },
  { id: "s13", name: "Guineo", type: "Materia Prima", stock: 36 },
  { id: "s14", name: "Fresas", type: "Materia Prima", stock: 0 },
  { id: "s15", name: "Piña", type: "Materia Prima", stock: 3 },
  { id: "s16", name: "Sal", type: "Materia Prima", stock: 0 },
  { id: "s17", name: "Limon", type: "Materia Prima", stock: 200 },
  { id: "s18", name: "Chile", type: "Materia Prima", stock: 1 },
  { id: "s19", name: "Aiuashte", type: "Materia Prima", stock: 1 },
  { id: "s20", name: "Salsa negra", type: "Materia Prima", stock: 1 },
  { id: "s21", name: "Tajin", type: "Materia Prima", stock: 0 },
  { id: "s22", name: "Chamoy", type: "Materia Prima", stock: 0 },
  { id: "s23", name: "Leche condensada", type: "Materia Prima", stock: 0 },
  { id: "s24", name: "Chocolate", type: "Materia Prima", stock: 0 },
  { id: "s25", name: "Picsi", type: "Materia Prima", stock: 0 },
  { id: "s26", name: "Crispy", type: "Materia Prima", stock: 0 },
  { id: "s27", name: "Nachos", type: "Materia Prima", stock: 72 },
  { id: "s28", name: "Jalapeños", type: "Materia Prima", stock: 0 },
  { id: "s29", name: "Hielo", type: "Materia Prima", stock: 400 },
  { id: "s30", name: "Gomita", type: "Materia Prima", stock: 0 },
  { id: "s31", name: "Mango", type: "Materia Prima", stock: 0 },
  { id: "s32", name: "Pepino", type: "Materia Prima", stock: 25 },
  { id: "s33", name: "Tomate", type: "Materia Prima", stock: 75 },
  { id: "s34", name: "Cebolla", type: "Materia Prima", stock: 0 },
  { id: "s35", name: "Cilantro", type: "Materia Prima", stock: 0 },
  { id: "s36", name: "Jicama", type: "Materia Prima", stock: 0 },
  { id: "s37", name: "Camaron (lote 1)", type: "Materia Prima", stock: 2000 },
  { id: "s38", name: "Camaron (lote 2)", type: "Materia Prima", stock: 2000 },
  {
    id: "s39",
    name: "agua alpina en botella",
    type: "Materia Prima",
    stock: 72,
  },
  { id: "s40", name: "lata sutida soda", type: "Materia Prima", stock: 24 },
  { id: "s41", name: "lata coca cola soda", type: "Materia Prima", stock: 48 },
];

const INITIAL_EMPLOYEES = [
  {
    id: "emp-kat",
    name: "Kat",
    startDate: "2026-12-03",
    payDates: [
      {
        id: "kat-1",
        label: "Pago 1 (Quincena)",
        dueDate: "2026-12-15",
        amount: 201.6,
        note: "13,44 x día",
        paid: false,
        paidDate: "",
      },
      {
        id: "kat-2",
        label: "Pago 2 (Fin de mes)",
        dueDate: "2026-12-30",
        amount: 201.6,
        note: "13,44 x día",
        paid: false,
        paidDate: "",
      },
    ],
  },
  {
    id: "emp-gab",
    name: "Gab",
    startDate: "2026-12-04",
    payDates: [
      {
        id: "gab-1",
        label: "Pago 1 (Quincena)",
        dueDate: "2026-12-15",
        amount: 275.0,
        note: "18,33 x día",
        paid: false,
        paidDate: "",
      },
      {
        id: "gab-2",
        label: "Pago 2 (Fin de mes)",
        dueDate: "2026-12-30",
        amount: 275.0,
        note: "18,33 x día",
        paid: false,
        paidDate: "",
      },
    ],
  },
];

const INITIAL_FINANCIAL = {
  goal: 7000.0,
  employees: INITIAL_EMPLOYEES,
  otherDebts: [
    {
      id: "d-prestamo",
      concept: "Prestamo",
      amount: 1000.0,
      date: "2026-12-03",
      paid: false,
    },
    {
      id: "d-intereses",
      concept: "Intereses",
      amount: 200.0,
      date: "2026-12-15",
      paid: false,
    },
    {
      id: "d-tarjeta",
      concept: "Plazo/tarjeta",
      amount: 696.9,
      date: "2026-12-30",
      paid: false,
    },
    {
      id: "d-ever",
      concept: "Contado Ever",
      amount: 0.0,
      date: "2026-12-03",
      paid: true,
    },
  ],
  reinvestments: [],
};

function syncProductsWithSharedSupplies(prodList, supList) {
  const traySup = supList.find((s) => s.id === "s3");
  const licuadoCupSup = supList.find((s) => s.id === "s1");

  return prodList.map((p) => {
    const recipe = Array.isArray(p.recipe) ? p.recipe : [];
    if (recipe.length > 0) {
      let maxByRecipe = Infinity;
      for (const r of recipe) {
        const sup = supList.find((s) => s.id === r.supplyId);
        const qtyPerUnit = parseFloat(r.qty) || 0;
        if (sup && qtyPerUnit > 0) {
          maxByRecipe = Math.min(
            maxByRecipe,
            Math.floor(sup.stock / qtyPerUnit)
          );
        }
      }
      if (Number.isFinite(maxByRecipe)) {
        return { ...p, recipe, stock: Math.max(0, maxByRecipe) };
      }
    }

    if (p.category === "Minutas" && traySup) {
      return { ...p, recipe, stock: traySup.stock };
    }
    if (p.category === "Licuados" && licuadoCupSup) {
      return { ...p, recipe, stock: licuadoCupSup.stock };
    }
    return { ...p, recipe };
  });
}

function applySupplyNetChanges(supList, netMinutasDelta, netLicuadosDelta) {
  const netTotalDrinks = netMinutasDelta + netLicuadosDelta;
  return supList.map((sup) => {
    let delta = 0;
    if (sup.id === "s3") delta += netMinutasDelta;
    if (sup.id === "s2") delta += netMinutasDelta;
    if (sup.id === "s1") delta += netLicuadosDelta;
    if (sup.id === "s8") delta += netTotalDrinks;

    if (delta !== 0) {
      return { ...sup, stock: Math.max(0, sup.stock - delta) };
    }
    return sup;
  });
}

function computeItemsSupplyImpact(itemsList, productsCatalog, sign = 1) {
  const recipeDeltas = {};
  let fallbackMinutas = 0;
  let fallbackLicuados = 0;

  for (const item of itemsList) {
    const qtySold = parseInt(item.qty, 10) || 0;
    if (qtySold === 0) continue;

    const catalogProd = productsCatalog.find((p) => p.id === item.id);
    const recipe = Array.isArray(catalogProd?.recipe)
      ? catalogProd.recipe
      : Array.isArray(item.recipe)
      ? item.recipe
      : [];

    if (recipe.length > 0) {
      for (const r of recipe) {
        const perUnit = parseFloat(r.qty) || 0;
        if (r.supplyId && perUnit > 0) {
          recipeDeltas[r.supplyId] =
            (recipeDeltas[r.supplyId] || 0) + sign * qtySold * perUnit;
        }
      }
    } else {
      const cat = catalogProd?.category || item.category;
      if (cat === "Minutas") fallbackMinutas += sign * qtySold;
      if (cat === "Licuados") fallbackLicuados += sign * qtySold;
    }
  }

  return { recipeDeltas, fallbackMinutas, fallbackLicuados };
}

function applyCombinedSupplyChanges(
  supList,
  recipeDeltas,
  fallbackMinutas,
  fallbackLicuados
) {
  const baseUpdated = applySupplyNetChanges(
    supList,
    fallbackMinutas,
    fallbackLicuados
  );
  return baseUpdated.map((sup) => {
    const extraDelta = recipeDeltas[sup.id] || 0;
    if (extraDelta !== 0) {
      return { ...sup, stock: Math.max(0, sup.stock - extraDelta) };
    }
    return sup;
  });
}

export default function NegocioElCheleApp() {
  const [activeTab, setActiveTab] = useState("pos");
  const [products, setProducts] = useState([]);
  const [supplies, setSupplies] = useState([]);
  const [inventoryLogs, setInventoryLogs] = useState([]);
  const [sales, setSales] = useState([]);
  const [financial, setFinancial] = useState(INITIAL_FINANCIAL);

  const [search, setSearch] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("Todos");

  // Carrito y cobro
  const [cart, setCart] = useState([]);
  const [paymentMethod, setPaymentMethod] = useState("Efectivo");
  const [cashGiven, setCashGiven] = useState("");
  const [transactionCode, setTransactionCode] = useState("");
  const [saleSuccess, setSaleSuccess] = useState(false);

  // Contraseñas: Precio/Stock, Financiero y Modificar/Eliminar Ventas
  const [priceAndStockUnlocked, setPriceAndStockUnlocked] = useState(false);
  const [financialUnlocked, setFinancialUnlocked] = useState(false);
  const [salesEditUnlocked, setSalesEditUnlocked] = useState(false);
  const [passwordInput, setPasswordInput] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [pendingAction, setPendingAction] = useState(null);
  const [salePendingAction, setSalePendingAction] = useState(null);
  const [productPendingEdit, setProductPendingEdit] = useState(null);

  // Modal para editar producto del menú y sus insumos vinculados
  const [editingProduct, setEditingProduct] = useState(null);
  const [editProdName, setEditProdName] = useState("");
  const [editProdCategory, setEditProdCategory] = useState("Minutas");
  const [editProdPrice, setEditProdPrice] = useState("");
  const [editProdStock, setEditProdStock] = useState("");
  const [editProdDesc, setEditProdDesc] = useState("");
  const [editProdRecipe, setEditProdRecipe] = useState([]);
  const [editSelectedSupplyId, setEditSelectedSupplyId] = useState("");
  const [editSelectedSupplyQty, setEditSelectedSupplyQty] = useState("1");

  // Modal para modificar una venta existente
  const [editingSale, setEditingSale] = useState(null);
  const [editSaleItems, setEditSaleItems] = useState([]);
  const [editSaleMethod, setEditSaleMethod] = useState("Efectivo");
  const [editSaleCashGiven, setEditSaleCashGiven] = useState("");
  const [editSaleTxCode, setEditSaleTxCode] = useState("");

  // Modal "Actualizar inventario"
  const [updateInvModalOpen, setUpdateInvModalOpen] = useState(false);
  const [invUpdateDate, setInvUpdateDate] = useState(
    new Date().toISOString().slice(0, 10)
  );
  const [invUpdateMode, setInvUpdateMode] = useState("entrada");
  const [invUpdateNote, setInvUpdateNote] = useState("");
  const [invDraftInputs, setInvDraftInputs] = useState({});

  // Crear nuevo insumo
  const [newSupplyName, setNewSupplyName] = useState("");
  const [newSupplyType, setNewSupplyType] = useState("Materia Prima");
  const [newSupplyStock, setNewSupplyStock] = useState("");

  // Nuevo producto de menú + vinculación de insumos
  const [newProdName, setNewProdName] = useState("");
  const [newProdCategory, setNewProdCategory] = useState("Minutas");
  const [newProdPrice, setNewProdPrice] = useState("");
  const [newProdStock, setNewProdStock] = useState("");
  const [newProdDesc, setNewProdDesc] = useState("");
  const [newProdRecipe, setNewProdRecipe] = useState([]);
  const [newSelectedSupplyId, setNewSelectedSupplyId] = useState("");
  const [newSelectedSupplyQty, setNewSelectedSupplyQty] = useState("1");

  // Nuevos registros financieros
  const [newDebtConcept, setNewDebtConcept] = useState("");
  const [newDebtAmount, setNewDebtAmount] = useState("");
  const [newDebtDate, setNewDebtDate] = useState("2026-12-15");

  const [newReinvConcept, setNewReinvConcept] = useState("");
  const [newReinvCategory, setNewReinvCategory] = useState("Minutas");
  const [newReinvAmount, setNewReinvAmount] = useState("");
  const [newReinvDate, setNewReinvDate] = useState("2026-12-04");

  const [extraPayEmpId, setExtraPayEmpId] = useState("emp-kat");
  const [extraPayLabel, setExtraPayLabel] = useState("Pago adicional");
  const [extraPayDate, setExtraPayDate] = useState("2026-12-15");
  const [extraPayAmount, setExtraPayAmount] = useState("");

  useEffect(() => {
    const savedProducts = localStorage.getItem("chele_v7_products");
    const savedSupplies = localStorage.getItem("chele_v7_supplies");
    const savedInvLogs = localStorage.getItem("chele_v7_inv_logs");
    const savedSales = localStorage.getItem("chele_v7_sales");
    const savedFinancial = localStorage.getItem("chele_v7_financial");

    const loadedSupplies = savedSupplies
      ? JSON.parse(savedSupplies)
      : INITIAL_SUPPLIES;
    const loadedProducts = savedProducts
      ? JSON.parse(savedProducts)
      : INITIAL_MENU;

    setSupplies(loadedSupplies);
    setProducts(syncProductsWithSharedSupplies(loadedProducts, loadedSupplies));
    setInventoryLogs(savedInvLogs ? JSON.parse(savedInvLogs) : []);
    setSales(savedSales ? JSON.parse(savedSales) : []);
    setFinancial(
      savedFinancial ? JSON.parse(savedFinancial) : INITIAL_FINANCIAL
    );

    if (loadedSupplies.length > 0) {
      setNewSelectedSupplyId((prev) => prev || loadedSupplies[0].id);
      setEditSelectedSupplyId((prev) => prev || loadedSupplies[0].id);
    }
  }, []);

  useEffect(() => {
    if (products.length > 0) {
      localStorage.setItem("chele_v7_products", JSON.stringify(products));
    }
  }, [products]);

  useEffect(() => {
    if (supplies.length > 0) {
      localStorage.setItem("chele_v7_supplies", JSON.stringify(supplies));
    }
  }, [supplies]);

  useEffect(() => {
    localStorage.setItem("chele_v7_inv_logs", JSON.stringify(inventoryLogs));
  }, [inventoryLogs]);

  useEffect(() => {
    localStorage.setItem("chele_v7_sales", JSON.stringify(sales));
  }, [sales]);

  useEffect(() => {
    localStorage.setItem("chele_v7_financial", JSON.stringify(financial));
  }, [financial]);

  // --- CONTRASEÑA ---
  const openPasswordModal = (target, payloadObj = null) => {
    setPasswordInput("");
    setPasswordError("");
    setPendingAction(target);
    if (target === "editProduct") {
      setProductPendingEdit(payloadObj);
    } else {
      setSalePendingAction(payloadObj);
    }
  };

  const openSaleEditorDirectly = (saleObj) => {
    setEditingSale(saleObj);
    setEditSaleItems(saleObj.items.map((it) => ({ ...it })));
    setEditSaleMethod(saleObj.paymentMethod);
    setEditSaleCashGiven(
      saleObj.paymentMethod === "Efectivo" ? String(saleObj.cashGiven ?? "") : ""
    );
    setEditSaleTxCode(saleObj.transactionCode || "");
  };

  const openProductEditorDirectly = (prodObj) => {
    setEditingProduct(prodObj);
    setEditProdName(prodObj.name || "");
    setEditProdCategory(prodObj.category || "Minutas");
    setEditProdPrice(String(prodObj.price ?? ""));
    setEditProdStock(String(prodObj.stock ?? ""));
    setEditProdDesc(prodObj.description || "");
    setEditProdRecipe(Array.isArray(prodObj.recipe) ? [...prodObj.recipe] : []);
    if (supplies.length > 0) {
      setEditSelectedSupplyId(supplies[0].id);
    }
    setEditSelectedSupplyQty("1");
  };

  const deleteSaleAndRestoreStock = (saleObj) => {
    const confirmed = window.confirm(
      `¿Eliminar la venta del ${saleObj.date} por $${Number(
        saleObj.total
      ).toFixed(2)} y devolver su stock?`
    );
    if (!confirmed) return;

    const { recipeDeltas, fallbackMinutas, fallbackLicuados } =
      computeItemsSupplyImpact(saleObj.items, products, -1);

    const updatedSupplies = applyCombinedSupplyChanges(
      supplies,
      recipeDeltas,
      fallbackMinutas,
      fallbackLicuados
    );
    setSupplies(updatedSupplies);

    const updatedProducts = syncProductsWithSharedSupplies(
      products.map((prod) => {
        const hasRecipe = Array.isArray(prod.recipe) && prod.recipe.length > 0;
        if (
          !hasRecipe &&
          (prod.category === "Minutas" || prod.category === "Licuados")
        ) {
          return prod;
        }
        const soldItem = saleObj.items.find((i) => i.id === prod.id);
        const qtyToRestore = parseInt(soldItem?.qty, 10) || 0;
        if (qtyToRestore > 0 && !hasRecipe) {
          return { ...prod, stock: prod.stock + qtyToRestore };
        }
        return prod;
      }),
      updatedSupplies
    );
    setProducts(updatedProducts);

    setSales(sales.filter((s) => s.id !== saleObj.id));
  };

  const handlePasswordSubmit = (e) => {
    e.preventDefault();
    if (passwordInput === OWNER_PASSWORD) {
      if (pendingAction === "priceAndStock") {
        setPriceAndStockUnlocked(true);
      } else if (pendingAction === "editProduct") {
        setPriceAndStockUnlocked(true);
        if (productPendingEdit) {
          openProductEditorDirectly(productPendingEdit);
        }
      } else if (pendingAction === "financial") {
        setFinancialUnlocked(true);
        setActiveTab("financial");
      } else if (pendingAction === "editSale") {
        setSalesEditUnlocked(true);
        if (salePendingAction) {
          openSaleEditorDirectly(salePendingAction);
        }
      } else if (pendingAction === "deleteSale") {
        setSalesEditUnlocked(true);
        const targetSale = salePendingAction;
        setPendingAction(null);
        setSalePendingAction(null);
        setPasswordInput("");
        setPasswordError("");
        if (targetSale) {
          deleteSaleAndRestoreStock(targetSale);
        }
        return;
      }
      setPendingAction(null);
      setSalePendingAction(null);
      setProductPendingEdit(null);
      setPasswordInput("");
      setPasswordError("");
    } else {
      setPasswordError("Contraseña incorrecta. Ingresa Negocio2026.");
    }
  };

  // --- CREAR Y ELIMINAR INSUMOS ---
  const handleAddSupply = (e) => {
    e.preventDefault();
    if (!newSupplyName.trim()) return;
    if (!priceAndStockUnlocked) {
      openPasswordModal("priceAndStock");
      return;
    }

    const parsedStock = parseFloat(newSupplyStock);
    const createdSupply = {
      id: `sup-${Date.now()}`,
      name: newSupplyName.trim(),
      type: newSupplyType,
      stock: Number.isNaN(parsedStock) ? 0 : Math.max(0, parsedStock),
    };

    const nextSupplies = [...supplies, createdSupply];
    setSupplies(nextSupplies);
    setProducts(syncProductsWithSharedSupplies(products, nextSupplies));

    if (!newSelectedSupplyId) setNewSelectedSupplyId(createdSupply.id);
    if (!editSelectedSupplyId) setEditSelectedSupplyId(createdSupply.id);

    setNewSupplyName("");
    setNewSupplyStock("");
  };

  const handleDeleteSupply = (supObj) => {
    if (!priceAndStockUnlocked) {
      openPasswordModal("priceAndStock");
      return;
    }

    const confirmed = window.confirm(
      `¿Eliminar el insumo "${supObj.name}" del inventario? También se quitará de las recetas de productos.`
    );
    if (!confirmed) return;

    const nextSupplies = supplies.filter((s) => s.id !== supObj.id);
    const nextProducts = products.map((p) => ({
      ...p,
      recipe: Array.isArray(p.recipe)
        ? p.recipe.filter((r) => r.supplyId !== supObj.id)
        : [],
    }));

    setSupplies(nextSupplies);
    setProducts(syncProductsWithSharedSupplies(nextProducts, nextSupplies));

    if (newSelectedSupplyId === supObj.id) {
      setNewSelectedSupplyId(nextSupplies[0]?.id || "");
    }
    if (editSelectedSupplyId === supObj.id) {
      setEditSelectedSupplyId(nextSupplies[0]?.id || "");
    }
  };

  // --- EDITAR PRODUCTO Y SUS INSUMOS VINCULADOS ---
  const handleRequestEditProduct = (prodObj) => {
    if (!priceAndStockUnlocked) {
      openPasswordModal("editProduct", prodObj);
      return;
    }
    openProductEditorDirectly(prodObj);
  };

  const addSupplyToRecipe = (recipeList, setRecipeList, supplyId, qtyRaw) => {
    const qty = parseFloat(qtyRaw);
    if (!supplyId || Number.isNaN(qty) || qty <= 0) return;

    const exists = recipeList.find((r) => r.supplyId === supplyId);
    if (exists) {
      setRecipeList(
        recipeList.map((r) =>
          r.supplyId === supplyId
            ? { ...r, qty: (parseFloat(r.qty) || 0) + qty }
            : r
        )
      );
    } else {
      setRecipeList([...recipeList, { supplyId, qty }]);
    }
  };

  const updateRecipeItemQty = (recipeList, setRecipeList, supplyId, qtyRaw) => {
    const qty = parseFloat(qtyRaw);
    setRecipeList(
      recipeList.map((r) =>
        r.supplyId === supplyId
          ? { ...r, qty: Number.isNaN(qty) ? 0 : Math.max(0, qty) }
          : r
      )
    );
  };

  const removeSupplyFromRecipe = (recipeList, setRecipeList, supplyId) => {
    setRecipeList(recipeList.filter((r) => r.supplyId !== supplyId));
  };

  const handleSaveEditedProduct = (e) => {
    e.preventDefault();
    if (!editingProduct || !editProdName.trim()) return;

    const cleanRecipe = editProdRecipe
      .map((r) => ({ supplyId: r.supplyId, qty: parseFloat(r.qty) || 0 }))
      .filter((r) => r.supplyId && r.qty > 0);

    const parsedPrice = parseFloat(editProdPrice);
    const parsedStock = parseInt(editProdStock, 10);

    const updatedList = products.map((p) => {
      if (p.id !== editingProduct.id) return p;
      return {
        ...p,
        name: editProdName.trim(),
        category: editProdCategory,
        price: Number.isNaN(parsedPrice) ? p.price : parsedPrice,
        stock: Number.isNaN(parsedStock) ? p.stock : Math.max(0, parsedStock),
        description: editProdDesc.trim(),
        recipe: cleanRecipe,
      };
    });

    setProducts(syncProductsWithSharedSupplies(updatedList, supplies));
    setEditingProduct(null);
  };

  // --- MODIFICAR / ELIMINAR VENTA EXISTENTE ---
  const handleRequestEditSale = (saleObj) => {
    if (!salesEditUnlocked) {
      openPasswordModal("editSale", saleObj);
      return;
    }
    openSaleEditorDirectly(saleObj);
  };

  const handleRequestDeleteSale = (saleObj) => {
    if (!salesEditUnlocked) {
      openPasswordModal("deleteSale", saleObj);
      return;
    }
    deleteSaleAndRestoreStock(saleObj);
  };

  const handleUpdateEditItemQty = (itemId, newQtyRaw) => {
    const qty = Math.max(0, parseInt(newQtyRaw, 10) || 0);
    setEditSaleItems(
      editSaleItems.map((it) => (it.id === itemId ? { ...it, qty } : it))
    );
  };

  const editedSaleTotal = editSaleItems.reduce(
    (sum, it) =>
      sum + (parseFloat(it.price) || 0) * (parseInt(it.qty, 10) || 0),
    0
  );
  const editedCashNum = parseFloat(editSaleCashGiven);
  const editedChangeAmount =
    editSaleMethod === "Efectivo" && !Number.isNaN(editedCashNum)
      ? editedCashNum - editedSaleTotal
      : 0;

  const handleSaveEditedSale = (e) => {
    e.preventDefault();
    if (!editingSale) return;

    const cleanedItems = editSaleItems
      .map((it) => ({ ...it, qty: Math.max(0, parseInt(it.qty, 10) || 0) }))
      .filter((it) => it.qty > 0);

    if (cleanedItems.length === 0) {
      alert("La venta debe tener al menos 1 producto con cantidad mayor a 0.");
      return;
    }

    if (
      editSaleMethod === "Efectivo" &&
      (Number.isNaN(editedCashNum) || editedCashNum < editedSaleTotal)
    ) {
      alert("El efectivo recibido debe ser igual o mayor al nuevo total.");
      return;
    }

    const oldImpact = computeItemsSupplyImpact(editingSale.items, products, 1);
    const newImpact = computeItemsSupplyImpact(cleanedItems, products, 1);

    const netRecipeDeltas = {};
    const allSupplyIds = new Set([
      ...Object.keys(oldImpact.recipeDeltas),
      ...Object.keys(newImpact.recipeDeltas),
    ]);
    for (const supId of allSupplyIds) {
      netRecipeDeltas[supId] =
        (newImpact.recipeDeltas[supId] || 0) -
        (oldImpact.recipeDeltas[supId] || 0);
    }

    const updatedSupplies = applyCombinedSupplyChanges(
      supplies,
      netRecipeDeltas,
      newImpact.fallbackMinutas - oldImpact.fallbackMinutas,
      newImpact.fallbackLicuados - oldImpact.fallbackLicuados
    );
    setSupplies(updatedSupplies);

    const updatedProducts = syncProductsWithSharedSupplies(
      products.map((prod) => {
        const hasRecipe = Array.isArray(prod.recipe) && prod.recipe.length > 0;
        if (
          !hasRecipe &&
          (prod.category === "Minutas" || prod.category === "Licuados")
        ) {
          return prod;
        }
        const prevItem = editingSale.items.find((i) => i.id === prod.id);
        const nextItem = cleanedItems.find((i) => i.id === prod.id);
        const diff = (nextItem?.qty || 0) - (prevItem?.qty || 0);
        return diff !== 0 && !hasRecipe
          ? { ...prod, stock: Math.max(0, prod.stock - diff) }
          : prod;
      }),
      updatedSupplies
    );
    setProducts(updatedProducts);

    const cleanTxCode =
      editSaleMethod === "Transferencia" || editSaleMethod === "Tarjeta"
        ? editSaleTxCode.trim()
        : "";

    const updatedSalesList = sales.map((s) =>
      s.id === editingSale.id
        ? {
            ...s,
            items: cleanedItems,
            total: editedSaleTotal,
            paymentMethod: editSaleMethod,
            transactionCode: cleanTxCode,
            cashGiven:
              editSaleMethod === "Efectivo" ? editedCashNum : editedSaleTotal,
            change:
              editSaleMethod === "Efectivo"
                ? Math.max(0, editedChangeAmount)
                : 0,
            editedAt: new Date().toLocaleString("es-SV"),
          }
        : s
    );

    setSales(updatedSalesList);
    setEditingSale(null);
    setEditSaleItems([]);
    setEditSaleCashGiven("");
    setEditSaleTxCode("");
  };

  // --- ACTUALIZAR INVENTARIO DE INSUMOS CON HISTORIAL ---
  const openUpdateInventoryModal = () => {
    const initialMap = {};
    supplies.forEach((s) => {
      initialMap[s.id] = "";
    });
    setInvDraftInputs(initialMap);
    setInvUpdateDate(new Date().toISOString().slice(0, 10));
    setInvUpdateMode("entrada");
    setInvUpdateNote("");
    setUpdateInvModalOpen(true);
  };

  const handleSaveInventoryUpdate = (e) => {
    e.preventDefault();
    const changedItems = [];

    const updatedSupplies = supplies.map((sup) => {
      const rawVal = invDraftInputs[sup.id];
      if (rawVal === "" || rawVal === undefined) return sup;

      const qty = parseFloat(rawVal);
      if (Number.isNaN(qty) || qty < 0) return sup;

      const prevStock = sup.stock;
      let nextStock = prevStock;
      let delta = 0;

      if (invUpdateMode === "entrada") {
        nextStock = prevStock + qty;
        delta = qty;
      } else if (invUpdateMode === "salida") {
        nextStock = Math.max(0, prevStock - qty);
        delta = nextStock - prevStock;
      } else if (invUpdateMode === "conteo") {
        nextStock = qty;
        delta = nextStock - prevStock;
      }

      if (nextStock !== prevStock) {
        changedItems.push({
          supplyId: sup.id,
          name: sup.name,
          type: sup.type,
          mode: invUpdateMode,
          inputQty: qty,
          previousStock: prevStock,
          newStock: nextStock,
          delta,
        });
      }

      return { ...sup, stock: nextStock };
    });

    if (changedItems.length === 0) {
      alert("Ingresa al menos una cantidad para actualizar el inventario.");
      return;
    }

    setSupplies(updatedSupplies);
    setProducts(syncProductsWithSharedSupplies(products, updatedSupplies));

    const newLogEntry = {
      id: `log-${Date.now()}`,
      date: invUpdateDate,
      recordedAt: new Date().toLocaleString("es-SV"),
      mode: invUpdateMode,
      note: invUpdateNote.trim(),
      items: changedItems,
    };

    setInventoryLogs([newLogEntry, ...inventoryLogs]);
    setUpdateInvModalOpen(false);
  };

  // --- VENTAS Y CARRITO ---
  const addToCart = (product) => {
    if (product.stock <= 0) return;

    const hasRecipe = Array.isArray(product.recipe) && product.recipe.length > 0;
    const sameCategoryCartQty = cart
      .filter((item) => item.category === product.category)
      .reduce((sum, item) => sum + item.qty, 0);

    const existing = cart.find((item) => item.id === product.id);
    const currentItemQty = existing ? existing.qty : 0;

    if (hasRecipe && currentItemQty >= product.stock) return;
    if (!hasRecipe && sameCategoryCartQty >= product.stock) return;

    if (existing) {
      setCart(
        cart.map((item) =>
          item.id === product.id ? { ...item, qty: item.qty + 1 } : item
        )
      );
    } else {
      setCart([...cart, { ...product, qty: 1 }]);
    }
  };

  const updateCartQty = (id, delta) => {
    const prod = products.find((p) => p.id === id);
    const hasRecipe = Array.isArray(prod?.recipe) && prod.recipe.length > 0;
    const sameCategoryCartQty = prod
      ? cart
          .filter((item) => item.category === prod.category)
          .reduce((sum, item) => sum + item.qty, 0)
      : 0;

    setCart(
      cart
        .map((item) => {
          if (item.id === id) {
            const newQty = item.qty + delta;
            if (newQty <= 0) return null;
            if (
              delta > 0 &&
              prod &&
              !hasRecipe &&
              (prod.category === "Minutas" || prod.category === "Licuados") &&
              sameCategoryCartQty >= prod.stock
            ) {
              return item;
            }
            if (prod && newQty > prod.stock) return item;
            return { ...item, qty: newQty };
          }
          return item;
        })
        .filter(Boolean)
    );
  };

  const cartTotal = cart.reduce((sum, item) => sum + item.price * item.qty, 0);
  const cashNum = parseFloat(cashGiven);
  const changeAmount =
    paymentMethod === "Efectivo" && !Number.isNaN(cashNum)
      ? cashNum - cartTotal
      : 0;

  const handleCompleteSale = () => {
    if (cart.length === 0) return;
    if (
      paymentMethod === "Efectivo" &&
      (Number.isNaN(cashNum) || cashNum < cartTotal)
    ) {
      alert("El monto en efectivo debe ser igual o mayor al total.");
      return;
    }

    const { recipeDeltas, fallbackMinutas, fallbackLicuados } =
      computeItemsSupplyImpact(cart, products, 1);

    const updatedSupplies = applyCombinedSupplyChanges(
      supplies,
      recipeDeltas,
      fallbackMinutas,
      fallbackLicuados
    );
    setSupplies(updatedSupplies);

    const updatedProducts = syncProductsWithSharedSupplies(
      products.map((prod) => {
        const hasRecipe = Array.isArray(prod.recipe) && prod.recipe.length > 0;
        if (
          !hasRecipe &&
          (prod.category === "Minutas" || prod.category === "Licuados")
        ) {
          return prod;
        }
        const sold = cart.find((c) => c.id === prod.id);
        if (sold && !hasRecipe) {
          return { ...prod, stock: Math.max(0, prod.stock - sold.qty) };
        }
        return prod;
      }),
      updatedSupplies
    );
    setProducts(updatedProducts);

    const cleanTxCode =
      paymentMethod === "Transferencia" || paymentMethod === "Tarjeta"
        ? transactionCode.trim()
        : "";

    const newSale = {
      id: Date.now(),
      date: new Date().toLocaleString("es-SV"),
      items: cart,
      total: cartTotal,
      paymentMethod,
      transactionCode: cleanTxCode,
      cashGiven: paymentMethod === "Efectivo" ? cashNum : cartTotal,
      change: paymentMethod === "Efectivo" ? Math.max(0, changeAmount) : 0,
    };

    setSales([newSale, ...sales]);
    setCart([]);
    setCashGiven("");
    setTransactionCode("");
    setSaleSuccess(true);
    setTimeout(() => setSaleSuccess(false), 3000);
  };

  // --- PRECIO Y STOCK INICIAL PROTEGIDOS ---
  const handleProtectedPriceChange = (productId, rawValue) => {
    if (!priceAndStockUnlocked) {
      openPasswordModal("priceAndStock");
      return;
    }
    const parsed = parseFloat(rawValue);
    setProducts(
      products.map((p) =>
        p.id === productId
          ? { ...p, price: Number.isNaN(parsed) ? 0 : parsed }
          : p
      )
    );
  };

  const handleProtectedStockChange = (productId, rawValue) => {
    if (!priceAndStockUnlocked) {
      openPasswordModal("priceAndStock");
      return;
    }
    const parsed = parseInt(rawValue, 10);
    const cleanVal = Number.isNaN(parsed) ? 0 : Math.max(0, parsed);
    const targetProd = products.find((p) => p.id === productId);

    if (!targetProd) return;

    const hasRecipe =
      Array.isArray(targetProd.recipe) && targetProd.recipe.length > 0;

    if (!hasRecipe && targetProd.category === "Minutas") {
      const nextSupplies = supplies.map((s) =>
        s.id === "s3" ? { ...s, stock: cleanVal } : s
      );
      setSupplies(nextSupplies);
      setProducts(syncProductsWithSharedSupplies(products, nextSupplies));
    } else if (!hasRecipe && targetProd.category === "Licuados") {
      const nextSupplies = supplies.map((s) =>
        s.id === "s1" ? { ...s, stock: cleanVal } : s
      );
      setSupplies(nextSupplies);
      setProducts(syncProductsWithSharedSupplies(products, nextSupplies));
    } else {
      setProducts(
        products.map((p) =>
          p.id === productId ? { ...p, stock: cleanVal } : p
        )
      );
    }
  };

  const handleAddProduct = (e) => {
    e.preventDefault();
    if (!newProdName.trim()) return;
    const priceVal = parseFloat(newProdPrice) || 0;

    if (
      !priceAndStockUnlocked &&
      (priceVal > 0 || newProdStock !== "" || newProdRecipe.length > 0)
    ) {
      openPasswordModal("priceAndStock");
      return;
    }

    const cleanRecipe = newProdRecipe
      .map((r) => ({ supplyId: r.supplyId, qty: parseFloat(r.qty) || 0 }))
      .filter((r) => r.supplyId && r.qty > 0);

    const trayStock = supplies.find((s) => s.id === "s3")?.stock ?? 400;
    const cupStock = supplies.find((s) => s.id === "s1")?.stock ?? 600;

    const initialStock =
      newProdStock !== ""
        ? Math.max(0, parseInt(newProdStock, 10) || 0)
        : newProdCategory === "Minutas"
        ? trayStock
        : cupStock;

    const created = {
      id: Date.now(),
      name: newProdName.trim(),
      category: newProdCategory,
      price: priceVal,
      stock: initialStock,
      description: newProdDesc.trim(),
      recipe: cleanRecipe,
    };

    setProducts(syncProductsWithSharedSupplies([...products, created], supplies));
    setNewProdName("");
    setNewProdPrice("");
    setNewProdStock("");
    setNewProdDesc("");
    setNewProdRecipe([]);
    setNewSelectedSupplyQty("1");
  };

  // --- FINANCIERO ---
  const toggleEmployeePayment = (empId, payId) => {
    setFinancial({
      ...financial,
      employees: financial.employees.map((emp) => {
        if (emp.id !== empId) return emp;
        return {
          ...emp,
          payDates: emp.payDates.map((pd) =>
            pd.id === payId
              ? {
                  ...pd,
                  paid: !pd.paid,
                  paidDate: !pd.paid
                    ? new Date().toISOString().slice(0, 10)
                    : "",
                }
              : pd
          ),
        };
      }),
    });
  };

  const handleAddEmployeePayDate = (e) => {
    e.preventDefault();
    const amt = parseFloat(extraPayAmount);
    if (!extraPayLabel.trim() || Number.isNaN(amt) || amt <= 0) return;

    setFinancial({
      ...financial,
      employees: financial.employees.map((emp) =>
        emp.id === extraPayEmpId
          ? {
              ...emp,
              payDates: [
                ...emp.payDates,
                {
                  id: `pay-${Date.now()}`,
                  label: extraPayLabel.trim(),
                  dueDate: extraPayDate,
                  amount: amt,
                  note: "Fecha registrada",
                  paid: false,
                  paidDate: "",
                },
              ],
            }
          : emp
      ),
    });
    setExtraPayLabel("Pago adicional");
    setExtraPayAmount("");
  };

  const handleAddDebt = (e) => {
    e.preventDefault();
    const amt = parseFloat(newDebtAmount);
    if (!newDebtConcept.trim() || Number.isNaN(amt) || amt <= 0) return;

    setFinancial({
      ...financial,
      otherDebts: [
        ...financial.otherDebts,
        {
          id: `debt-${Date.now()}`,
          concept: newDebtConcept.trim(),
          amount: amt,
          date: newDebtDate,
          paid: false,
        },
      ],
    });
    setNewDebtConcept("");
    setNewDebtAmount("");
  };

  const toggleDebtPaid = (debtId) => {
    setFinancial({
      ...financial,
      otherDebts: financial.otherDebts.map((d) =>
        d.id === debtId ? { ...d, paid: !d.paid } : d
      ),
    });
  };

  const handleAddReinvestment = (e) => {
    e.preventDefault();
    const amt = parseFloat(newReinvAmount);
    if (!newReinvConcept.trim() || Number.isNaN(amt) || amt <= 0) return;

    setFinancial({
      ...financial,
      reinvestments: [
        {
          id: `reinv-${Date.now()}`,
          concept: newReinvConcept.trim(),
          category: newReinvCategory,
          amount: amt,
          date: newReinvDate,
        },
        ...financial.reinvestments,
      ],
    });
    setNewReinvConcept("");
    setNewReinvAmount("");
  };

  // --- EXPORTAR EXCEL ---
  const exportToExcel = () => {
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        products.map((p) => ({
          Producto: p.name,
          Categoria: p.category,
          Precio: p.price,
          Stock: p.stock,
          Insumos_Vinculados: (p.recipe || [])
            .map((r) => {
              const sName =
                supplies.find((s) => s.id === r.supplyId)?.name || r.supplyId;
              return `${r.qty}x ${sName}`;
            })
            .join(", "),
        }))
      ),
      "Menu"
    );

    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        supplies.map((s) => ({
          Detalle: s.name,
          Tipo: s.type,
          "Total disponible": s.stock,
        }))
      ),
      "Inventario"
    );

    const flatLogs = inventoryLogs.flatMap((log) =>
      log.items.map((it) => ({
        Fecha_Movimiento: log.date,
        Registrado: log.recordedAt,
        Tipo_Actualizacion: log.mode,
        Insumo: it.name,
        Categoria: it.type,
        Cantidad_Ingresada: it.inputQty,
        Stock_Anterior: it.previousStock,
        Stock_Nuevo: it.newStock,
        Nota: log.note || "",
      }))
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(flatLogs),
      "Historial_Inventario"
    );

    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        sales.map((s) => ({
          Fecha: s.date,
          Productos: s.items.map((i) => `${i.qty}x ${i.name}`).join(", "),
          Total: Number(s.total.toFixed(2)),
          Metodo_Pago: s.paymentMethod,
          Codigo_Transaccion: s.transactionCode || "",
          Cambio: Number(s.change.toFixed(2)),
        }))
      ),
      "Ventas"
    );

    XLSX.writeFile(wb, "Minutas_el_Chele_Completo.xlsx");
  };

  // --- CÁLCULOS FINANCIEROS ---
  const totalSold = sales.reduce((acc, s) => acc + s.total, 0);

  const totalCashSales = sales
    .filter((s) => s.paymentMethod === "Efectivo")
    .reduce((acc, s) => acc + s.total, 0);

  const totalDigitalSales = totalSold - totalCashSales;

  const totalReinvestments = financial.reinvestments.reduce(
    (sum, r) => sum + (parseFloat(r.amount) || 0),
    0
  );

  const pendingEmployeePay = financial.employees.reduce((sum, emp) => {
    return (
      sum +
      emp.payDates
        .filter((pd) => !pd.paid)
        .reduce((s, pd) => s + (parseFloat(pd.amount) || 0), 0)
    );
  }, 0);

  const paidEmployeePay = financial.employees.reduce((sum, emp) => {
    return (
      sum +
      emp.payDates
        .filter((pd) => pd.paid)
        .reduce((s, pd) => s + (parseFloat(pd.amount) || 0), 0)
    );
  }, 0);

  const pendingOtherDebts = financial.otherDebts
    .filter((d) => !d.paid)
    .reduce((sum, d) => sum + (parseFloat(d.amount) || 0), 0);

  const paidOtherDebts = financial.otherDebts
    .filter((d) => d.paid)
    .reduce((sum, d) => sum + (parseFloat(d.amount) || 0), 0);

  const totalPaymentsMade = paidEmployeePay + paidOtherDebts;
  const totalPendingDebtsAndSalaries = pendingEmployeePay + pendingOtherDebts;

  const netCashInDrawer =
    totalCashSales - totalReinvestments - totalPaymentsMade;

  const availableCash = totalSold - totalReinvestments - totalPaymentsMade;

  const goalTarget = parseFloat(financial.goal) || 7000;
  const netTowardGoal =
    totalSold -
    totalPendingDebtsAndSalaries -
    totalReinvestments -
    totalPaymentsMade;
  const remainingForGoal = Math.max(0, goalTarget - netTowardGoal);
  const goalProgressPct = Math.min(
    100,
    Math.max(0, (netTowardGoal / (goalTarget || 1)) * 100)
  );

  const filteredProducts = products.filter((p) => {
    const matchesSearch = p.name.toLowerCase().includes(search.toLowerCase());
    const matchesCat =
      categoryFilter === "Todos" || p.category === categoryFilter;
    return matchesSearch && matchesCat;
  });

  const getSupplyName = (id) =>
    supplies.find((s) => s.id === id)?.name || id;

  return (
    <div className="min-h-screen flex flex-col bg-slate-100 text-slate-800">
      {/* MODAL DE CONTRASEÑA */}
      {pendingAction && (
        <div className="fixed inset-0 z-50 bg-black/60 flex items-center justify-center p-4">
          <form
            onSubmit={handlePasswordSubmit}
            className="bg-white rounded-2xl shadow-xl border border-slate-200 max-w-sm w-full p-6 space-y-4"
          >
            <div className="flex items-center gap-2 text-blue-800 font-bold text-lg">
              <Lock size={20} />
              <span>
                {pendingAction === "financial"
                  ? "Acceso a Financiero"
                  : pendingAction === "deleteSale"
                  ? "Desbloquear Eliminación de Ventas"
                  : pendingAction === "editSale"
                  ? "Desbloquear Edición de Ventas"
                  : "Desbloquear Edición de Menú y Stock"}
              </span>
            </div>
            <p className="text-sm text-slate-600">
              {pendingAction === "financial"
                ? "Ingresa la contraseña del dueño para acceder al panel financiero."
                : pendingAction === "deleteSale"
                ? "Ingresa la contraseña del dueño para eliminar ventas y devolver su stock."
                : pendingAction === "editSale"
                ? "Ingresa la contraseña del dueño para modificar ventas registradas."
                : "Ingresa la contraseña del dueño para modificar productos, insumos, precios o stock."}
            </p>
            <input
              type="password"
              autoFocus
              placeholder="Contraseña"
              value={passwordInput}
              onChange={(e) => setPasswordInput(e.target.value)}
              className="w-full p-2.5 border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-600"
            />
            {passwordError && (
              <p className="text-xs text-red-600 font-semibold">
                {passwordError}
              </p>
            )}
            <div className="flex justify-end gap-2 pt-1">
              <button
                type="button"
                onClick={() => {
                  setPendingAction(null);
                  setSalePendingAction(null);
                  setProductPendingEdit(null);
                }}
                className="px-4 py-2 rounded-xl border border-slate-300 text-sm font-medium hover:bg-slate-50"
              >
                Cancelar
              </button>
              <button
                type="submit"
                className="px-4 py-2 rounded-xl bg-blue-800 text-white text-sm font-semibold hover:bg-blue-900"
              >
                Entrar
              </button>
            </div>
          </form>
        </div>
      )}

      {/* MODAL PARA EDITAR PRODUCTO Y SUS INSUMOS VINCULADOS */}
      {editingProduct && (
        <div className="fixed inset-0 z-50 bg-black/60 flex items-center justify-center p-4">
          <form
            onSubmit={handleSaveEditedProduct}
            className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-2xl w-full p-6 space-y-4 max-h-[90vh] overflow-y-auto"
          >
            <div className="flex justify-between items-center border-b pb-3">
              <div>
                <h3 className="text-lg font-bold text-slate-900">
                  Editar Producto e Insumos Vinculados
                </h3>
                <p className="text-xs text-slate-500">
                  Define qué insumos ocupa por cada 1 unidad vendida de este
                  producto.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setEditingProduct(null)}
                className="p-1.5 rounded-lg hover:bg-slate-100"
              >
                <X size={20} />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-bold uppercase text-slate-500 mb-1">
                  Nombre del producto
                </label>
                <input
                  type="text"
                  required
                  value={editProdName}
                  onChange={(e) => setEditProdName(e.target.value)}
                  className="w-full p-2.5 border border-slate-300 rounded-xl bg-white text-sm font-semibold"
                />
              </div>

              <div>
                <label className="block text-xs font-bold uppercase text-slate-500 mb-1">
                  Categoría
                </label>
                <select
                  value={editProdCategory}
                  onChange={(e) => setEditProdCategory(e.target.value)}
                  className="w-full p-2.5 border border-slate-300 rounded-xl bg-white text-sm font-semibold"
                >
                  <option value="Minutas">Minutas</option>
                  <option value="Licuados">Licuados</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold uppercase text-slate-500 mb-1">
                  Precio ($)
                </label>
                <input
                  type="number"
                  step="0.01"
                  required
                  value={editProdPrice}
                  onChange={(e) => setEditProdPrice(e.target.value)}
                  className="w-full p-2.5 border border-slate-300 rounded-xl bg-white text-sm font-bold"
                />
              </div>

              <div>
                <label className="block text-xs font-bold uppercase text-slate-500 mb-1">
                  Stock / Disponibilidad
                </label>
                <input
                  type="number"
                  value={editProdStock}
                  onChange={(e) => setEditProdStock(e.target.value)}
                  className="w-full p-2.5 border border-slate-300 rounded-xl bg-white text-sm font-bold"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold uppercase text-slate-500 mb-1">
                Descripción (opcional)
              </label>
              <input
                type="text"
                value={editProdDesc}
                onChange={(e) => setEditProdDesc(e.target.value)}
                placeholder="Ingredientes o detalle del producto"
                className="w-full p-2.5 border border-slate-300 rounded-xl bg-white text-sm"
              />
            </div>

            <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold uppercase text-slate-700 flex items-center gap-1.5">
                  <Layers size={15} className="text-blue-700" /> Insumos que
                  ocupa este producto (por unidad)
                </span>
              </div>

              <div className="flex flex-wrap gap-2 items-end">
                <div className="flex-1 min-w-[180px]">
                  <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                    Insumo
                  </label>
                  <select
                    value={editSelectedSupplyId}
                    onChange={(e) => setEditSelectedSupplyId(e.target.value)}
                    className="w-full p-2 border rounded-lg bg-white text-sm"
                  >
                    {supplies.map((s) => (
                      <option key={s.id} value={s.id}>
                        {s.name} (Disp: {s.stock})
                      </option>
                    ))}
                  </select>
                </div>

                <div className="w-28">
                  <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                    Cantidad / ud
                  </label>
                  <input
                    type="number"
                    step="any"
                    min="0.01"
                    value={editSelectedSupplyQty}
                    onChange={(e) => setEditSelectedSupplyQty(e.target.value)}
                    className="w-full p-2 border rounded-lg bg-white text-sm font-semibold"
                  />
                </div>

                <button
                  type="button"
                  onClick={() =>
                    addSupplyToRecipe(
                      editProdRecipe,
                      setEditProdRecipe,
                      editSelectedSupplyId,
                      editSelectedSupplyQty
                    )
                  }
                  className="bg-blue-700 hover:bg-blue-800 text-white px-3.5 py-2 rounded-lg text-xs font-bold flex items-center gap-1"
                >
                  <Plus size={15} /> Vincular Insumo
                </button>
              </div>

              {editProdRecipe.length === 0 ? (
                <p className="text-xs text-slate-500">
                  Sin insumos personalizados: al vender descontará la regla base
                  de categoría (bandeja/vaso y pajilla).
                </p>
              ) : (
                <div className="space-y-2">
                  {editProdRecipe.map((r) => (
                    <div
                      key={r.supplyId}
                      className="flex justify-between items-center gap-2 bg-white p-2.5 rounded-lg border border-slate-200 text-sm"
                    >
                      <span className="font-semibold text-slate-800">
                        {getSupplyName(r.supplyId)}
                      </span>
                      <div className="flex items-center gap-2">
                        <span className="text-xs text-slate-500">Ocupa:</span>
                        <input
                          type="number"
                          step="any"
                          min="0"
                          value={r.qty}
                          onChange={(e) =>
                            updateRecipeItemQty(
                              editProdRecipe,
                              setEditProdRecipe,
                              r.supplyId,
                              e.target.value
                            )
                          }
                          className="w-20 p-1 border border-slate-300 rounded text-right text-xs font-bold"
                        />
                        <button
                          type="button"
                          onClick={() =>
                            removeSupplyFromRecipe(
                              editProdRecipe,
                              setEditProdRecipe,
                              r.supplyId
                            )
                          }
                          className="text-red-600 hover:bg-red-50 p-1 rounded"
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setEditingProduct(null)}
                className="px-4 py-2 rounded-xl border border-slate-300 text-sm font-medium hover:bg-slate-50"
              >
                Cancelar
              </button>
              <button
                type="submit"
                className="px-5 py-2 rounded-xl bg-blue-800 hover:bg-blue-900 text-white text-sm font-bold"
              >
                Guardar Producto
              </button>
            </div>
          </form>
        </div>
      )}

      {/* MODAL PARA MODIFICAR UNA VENTA */}
      {editingSale && (
        <div className="fixed inset-0 z-50 bg-black/60 flex items-center justify-center p-4">
          <form
            onSubmit={handleSaveEditedSale}
            className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-xl w-full p-6 space-y-4 max-h-[90vh] overflow-y-auto"
          >
            <div className="flex justify-between items-center border-b pb-3">
              <div>
                <h3 className="text-lg font-bold text-slate-900">
                  Modificar Venta
                </h3>
                <p className="text-xs text-slate-500">
                  Registrada el {editingSale.date}
                </p>
              </div>
              <button
                type="button"
                onClick={() => setEditingSale(null)}
                className="p-1.5 rounded-lg hover:bg-slate-100"
              >
                <X size={20} />
              </button>
            </div>

            <div className="space-y-2.5">
              <label className="block text-xs font-bold uppercase text-slate-500">
                Cantidades por producto (0 elimina el producto de la venta)
              </label>
              {editSaleItems.map((it) => (
                <div
                  key={it.id}
                  className="flex justify-between items-center gap-3 bg-slate-50 p-3 rounded-xl border border-slate-200"
                >
                  <div className="flex-1">
                    <p className="font-semibold text-sm text-slate-800">
                      {it.name}
                    </p>
                    <p className="text-xs text-slate-500">
                      ${Number(it.price).toFixed(2)} c/u · {it.category}
                    </p>
                  </div>
                  <input
                    type="number"
                    min="0"
                    value={it.qty}
                    onChange={(e) =>
                      handleUpdateEditItemQty(it.id, e.target.value)
                    }
                    className="w-20 p-1.5 border border-slate-300 rounded-lg bg-white text-right font-bold"
                  />
                  <div className="w-20 text-right font-bold text-sm text-slate-900">
                    ${(Number(it.price) * (parseInt(it.qty, 10) || 0)).toFixed(2)}
                  </div>
                </div>
              ))}
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2 border-t">
              <div>
                <label className="block text-xs font-bold uppercase text-slate-500 mb-1">
                  Método de Pago
                </label>
                <select
                  value={editSaleMethod}
                  onChange={(e) => setEditSaleMethod(e.target.value)}
                  className="w-full p-2.5 border border-slate-300 rounded-xl bg-white text-sm font-semibold"
                >
                  <option value="Efectivo">Efectivo</option>
                  <option value="Transferencia">Transferencia</option>
                  <option value="Tarjeta">Tarjeta</option>
                </select>
              </div>

              {editSaleMethod === "Efectivo" && (
                <div>
                  <label className="block text-xs font-bold uppercase text-slate-500 mb-1">
                    Efectivo recibido ($)
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    value={editSaleCashGiven}
                    onChange={(e) => setEditSaleCashGiven(e.target.value)}
                    className="w-full p-2.5 border border-slate-300 rounded-xl bg-white text-sm font-bold text-right"
                  />
                </div>
              )}

              {(editSaleMethod === "Transferencia" ||
                editSaleMethod === "Tarjeta") && (
                <div>
                  <label className="block text-xs font-bold uppercase text-slate-500 mb-1">
                    Código de transacción
                  </label>
                  <input
                    type="text"
                    placeholder="Ej. TXN-948271"
                    value={editSaleTxCode}
                    onChange={(e) => setEditSaleTxCode(e.target.value)}
                    className="w-full p-2.5 border border-slate-300 rounded-xl bg-white text-sm font-semibold"
                  />
                </div>
              )}
            </div>

            <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-200 flex justify-between items-center">
              <span className="font-bold text-slate-700">Nuevo Total:</span>
              <span className="text-xl font-extrabold text-blue-800">
                ${editedSaleTotal.toFixed(2)}
              </span>
            </div>

            {editSaleMethod === "Efectivo" && (
              <div className="flex justify-between items-center text-sm px-1">
                <span className="text-slate-600 font-medium">Cambio:</span>
                <span className="font-extrabold text-green-700">
                  ${editedChangeAmount > 0 ? editedChangeAmount.toFixed(2) : "0.00"}
                </span>
              </div>
            )}

            <div className="flex justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setEditingSale(null)}
                className="px-4 py-2 rounded-xl border border-slate-300 text-sm font-medium hover:bg-slate-50"
              >
                Cancelar
              </button>
              <button
                type="submit"
                className="px-5 py-2 rounded-xl bg-blue-800 hover:bg-blue-900 text-white text-sm font-bold"
              >
                Guardar Cambios
              </button>
            </div>
          </form>
        </div>
      )}

      {/* MODAL ACTUALIZAR INVENTARIO */}
      {updateInvModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/60 flex items-center justify-center p-4">
          <form
            onSubmit={handleSaveInventoryUpdate}
            className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-3xl w-full max-h-[90vh] flex flex-col"
          >
            <div className="p-5 border-b flex justify-between items-center">
              <div>
                <h3 className="text-lg font-bold text-slate-900">
                  Actualizar Inventario de Materiales e Insumos
                </h3>
                <p className="text-xs text-slate-500">
                  Registra cantidades y se guardará automáticamente en el
                  historial.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setUpdateInvModalOpen(false)}
                className="p-1.5 rounded-lg hover:bg-slate-100"
              >
                <X size={20} />
              </button>
            </div>

            <div className="p-5 border-b bg-slate-50 grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Fecha del registro
                </label>
                <input
                  type="date"
                  required
                  value={invUpdateDate}
                  onChange={(e) => setInvUpdateDate(e.target.value)}
                  className="w-full p-2 border rounded-xl bg-white text-sm"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Tipo de actualización
                </label>
                <select
                  value={invUpdateMode}
                  onChange={(e) => setInvUpdateMode(e.target.value)}
                  className="w-full p-2 border rounded-xl bg-white text-sm font-medium"
                >
                  <option value="entrada">Sumar entrada (+)</option>
                  <option value="salida">Restar salida (-)</option>
                  <option value="conteo">Fijar stock exacto (=)</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Nota / Referencia
                </label>
                <input
                  type="text"
                  placeholder="Ej. Compra de insumos del día"
                  value={invUpdateNote}
                  onChange={(e) => setInvUpdateNote(e.target.value)}
                  className="w-full p-2 border rounded-xl bg-white text-sm"
                />
              </div>
            </div>

            <div className="p-5 overflow-y-auto flex-1">
              <table className="w-full text-left border-collapse text-sm">
                <thead className="bg-slate-100 border-b text-xs uppercase text-slate-600 sticky top-0">
                  <tr>
                    <th className="p-2.5">Material / Insumo</th>
                    <th className="p-2.5">Tipo</th>
                    <th className="p-2.5">Stock actual</th>
                    <th className="p-2.5">
                      {invUpdateMode === "entrada"
                        ? "Cantidad a sumar"
                        : invUpdateMode === "salida"
                        ? "Cantidad a restar"
                        : "Nuevo stock"}
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y">
                  {supplies.map((sup) => (
                    <tr key={sup.id} className="hover:bg-slate-50">
                      <td className="p-2.5 font-medium">{sup.name}</td>
                      <td className="p-2.5 text-slate-500 text-xs">
                        {sup.type}
                      </td>
                      <td className="p-2.5 font-bold text-slate-700">
                        {sup.stock}
                      </td>
                      <td className="p-2.5">
                        <input
                          type="number"
                          step="any"
                          min="0"
                          placeholder="—"
                          value={invDraftInputs[sup.id] ?? ""}
                          onChange={(e) =>
                            setInvDraftInputs({
                              ...invDraftInputs,
                              [sup.id]: e.target.value,
                            })
                          }
                          className="w-28 p-1.5 border border-slate-300 rounded-lg bg-white font-semibold"
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="p-4 border-t bg-slate-50 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setUpdateInvModalOpen(false)}
                className="px-4 py-2 rounded-xl border border-slate-300 text-sm font-medium hover:bg-white"
              >
                Cancelar
              </button>
              <button
                type="submit"
                className="px-5 py-2 rounded-xl bg-green-600 hover:bg-green-700 text-white text-sm font-bold"
              >
                Guardar Actualización
              </button>
            </div>
          </form>
        </div>
      )}

      {/* HEADER */}
      <header className="bg-blue-800 text-white shadow-md">
        <div className="max-w-7xl mx-auto px-4 py-4 flex flex-wrap justify-between items-center gap-3">
          <div>
            <h1 className="text-xl sm:text-2xl font-bold tracking-tight">
              🍧 Minutas el Chele
            </h1>
            <p className="text-xs text-blue-200">
              Ventas, Inventario Diario y Control Financiero
            </p>
          </div>

          <nav className="flex flex-wrap gap-2">
            <button
              onClick={() => setActiveTab("pos")}
              className={`flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-sm font-semibold transition ${
                activeTab === "pos"
                  ? "bg-white text-blue-800 shadow"
                  : "bg-blue-900/70 hover:bg-blue-700"
              }`}
            >
              <ShoppingCart size={17} /> Ventas
            </button>

            <button
              onClick={() => setActiveTab("inventory")}
              className={`flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-sm font-semibold transition ${
                activeTab === "inventory"
                  ? "bg-white text-blue-800 shadow"
                  : "bg-blue-900/70 hover:bg-blue-700"
              }`}
            >
              <Package size={17} /> Inventario
            </button>

            <button
              onClick={() => setActiveTab("history")}
              className={`flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-sm font-semibold transition ${
                activeTab === "history"
                  ? "bg-white text-blue-800 shadow"
                  : "bg-blue-900/70 hover:bg-blue-700"
              }`}
            >
              <History size={17} /> Historial Ventas ({sales.length})
            </button>

            <button
              onClick={() =>
                financialUnlocked
                  ? setActiveTab("financial")
                  : openPasswordModal("financial")
              }
              className={`flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-sm font-semibold transition ${
                activeTab === "financial"
                  ? "bg-amber-400 text-slate-950 shadow"
                  : "bg-blue-900/70 hover:bg-blue-700 text-amber-200"
              }`}
            >
              {financialUnlocked ? <Unlock size={16} /> : <Lock size={16} />}
              Financiero
            </button>
          </nav>
        </div>
      </header>

      {/* MAIN */}
      <main className="max-w-7xl w-full mx-auto p-4 flex-1">
        {/* 1. VENTAS */}
        {activeTab === "pos" && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <div className="lg:col-span-2 space-y-4">
              <div className="flex flex-wrap gap-3 items-center justify-between">
                <div className="relative flex-1 min-w-[220px]">
                  <Search
                    className="absolute left-3 top-3 text-slate-400"
                    size={19}
                  />
                  <input
                    type="text"
                    placeholder="Buscar producto..."
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    className="w-full pl-10 pr-4 py-2.5 rounded-xl border border-slate-300 bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
                  />
                </div>

                <div className="flex gap-1.5">
                  {["Todos", "Minutas", "Licuados"].map((cat) => (
                    <button
                      key={cat}
                      onClick={() => setCategoryFilter(cat)}
                      className={`px-3 py-2 rounded-xl text-xs font-semibold border transition ${
                        categoryFilter === cat
                          ? "bg-blue-700 text-white border-blue-700"
                          : "bg-white text-slate-700 border-slate-300 hover:bg-slate-50"
                      }`}
                    >
                      {cat}
                    </button>
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
                {filteredProducts.map((product) => (
                  <button
                    key={product.id}
                    onClick={() => addToCart(product)}
                    disabled={product.stock <= 0}
                    className={`p-4 rounded-2xl border text-left transition flex flex-col justify-between min-h-[135px] ${
                      product.stock > 0
                        ? "bg-white border-slate-200 hover:border-blue-600 hover:shadow-md active:scale-[0.99]"
                        : "bg-slate-200 border-slate-300 opacity-60 cursor-not-allowed"
                    }`}
                  >
                    <div>
                      <span className="text-[11px] font-semibold uppercase tracking-wide text-blue-700 bg-blue-50 px-2 py-0.5 rounded-full">
                        {product.category}
                      </span>
                      <h3 className="font-bold text-slate-900 mt-1.5">
                        {product.name}
                      </h3>
                      {product.description && (
                        <p className="text-xs text-slate-500 line-clamp-2 mt-1">
                          {product.description}
                        </p>
                      )}
                    </div>

                    <div className="flex justify-between items-center w-full pt-3 mt-2 border-t border-slate-100">
                      <span className="text-lg font-extrabold text-blue-800">
                        ${product.price.toFixed(2)}
                      </span>
                      <span
                        className={`text-xs px-2 py-0.5 rounded-full font-semibold ${
                          product.stock > 20
                            ? "bg-green-100 text-green-800"
                            : product.stock > 0
                            ? "bg-amber-100 text-amber-800"
                            : "bg-red-100 text-red-800"
                        }`}
                      >
                        {product.stock > 0
                          ? `Stock: ${product.stock}`
                          : "Agotado"}
                      </span>
                    </div>
                  </button>
                ))}
              </div>
            </div>

            {/* CAJA */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-col justify-between h-fit sticky top-4">
              <div>
                <div className="flex items-center justify-between border-b pb-3 mb-3">
                  <h2 className="text-lg font-bold">Orden de Venta</h2>
                  {cart.length > 0 && (
                    <button
                      onClick={() => setCart([])}
                      className="text-xs text-red-600 font-medium hover:underline"
                    >
                      Limpiar
                    </button>
                  )}
                </div>

                {cart.length === 0 ? (
                  <p className="text-slate-400 text-center py-8 text-sm">
                    Selecciona productos para registrar una venta.
                  </p>
                ) : (
                  <div className="space-y-2.5 max-h-60 overflow-y-auto pr-1 mb-4">
                    {cart.map((item) => (
                      <div
                        key={item.id}
                        className="flex justify-between items-center bg-slate-50 p-2.5 rounded-xl border border-slate-100"
                      >
                        <div className="flex-1 pr-2">
                          <p className="font-semibold text-sm text-slate-800">
                            {item.name}
                          </p>
                          <p className="text-xs text-slate-500">
                            ${item.price.toFixed(2)} c/u
                          </p>
                        </div>

                        <div className="flex items-center gap-1.5">
                          <button
                            onClick={() => updateCartQty(item.id, -1)}
                            className="p-1 bg-slate-200 rounded-lg hover:bg-slate-300"
                          >
                            <Minus size={14} />
                          </button>
                          <span className="font-bold w-6 text-center text-sm">
                            {item.qty}
                          </span>
                          <button
                            onClick={() => updateCartQty(item.id, 1)}
                            className="p-1 bg-slate-200 rounded-lg hover:bg-slate-300"
                          >
                            <Plus size={14} />
                          </button>
                        </div>

                        <div className="w-16 text-right font-bold text-sm text-slate-900">
                          ${(item.price * item.qty).toFixed(2)}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div className="border-t pt-4 space-y-4">
                <div className="flex justify-between items-center text-xl font-extrabold">
                  <span>Total:</span>
                  <span className="text-blue-800">${cartTotal.toFixed(2)}</span>
                </div>

                <div>
                  <label className="block text-xs font-bold uppercase text-slate-500 mb-2">
                    Método de Pago
                  </label>
                  <div className="grid grid-cols-3 gap-2">
                    {[
                      { id: "Efectivo", icon: DollarSign },
                      { id: "Transferencia", icon: ArrowRightLeft },
                      { id: "Tarjeta", icon: CreditCard },
                    ].map(({ id, icon: Icon }) => (
                      <button
                        key={id}
                        type="button"
                        onClick={() => setPaymentMethod(id)}
                        className={`flex flex-col items-center justify-center py-2 px-1 rounded-xl border text-xs font-semibold transition ${
                          paymentMethod === id
                            ? "bg-blue-50 border-blue-700 text-blue-800"
                            : "border-slate-200 text-slate-600 hover:bg-slate-50"
                        }`}
                      >
                        <Icon size={16} className="mb-1" />
                        {id}
                      </button>
                    ))}
                  </div>
                </div>

                {paymentMethod === "Efectivo" && (
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-200 space-y-2">
                    <div className="flex justify-between items-center gap-2">
                      <label className="text-sm font-medium text-slate-700">
                        Efectivo recibido ($):
                      </label>
                      <input
                        type="number"
                        step="0.01"
                        placeholder="0.00"
                        value={cashGiven}
                        onChange={(e) => setCashGiven(e.target.value)}
                        className="w-28 text-right font-bold p-1.5 rounded-lg border border-slate-300 bg-white"
                      />
                    </div>
                    <div className="flex justify-between items-center pt-1.5 border-t border-slate-200">
                      <span className="text-sm font-medium text-slate-600">
                        Cambio:
                      </span>
                      <span
                        className={`text-lg font-extrabold ${
                          changeAmount < 0 ? "text-red-600" : "text-green-700"
                        }`}
                      >
                        ${changeAmount > 0 ? changeAmount.toFixed(2) : "0.00"}
                      </span>
                    </div>
                  </div>
                )}

                {(paymentMethod === "Transferencia" ||
                  paymentMethod === "Tarjeta") && (
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-200 space-y-1.5">
                    <label className="block text-xs font-bold uppercase text-slate-600">
                      Código de transacción ({paymentMethod})
                    </label>
                    <input
                      type="text"
                      placeholder="Ingresa el código o referencia"
                      value={transactionCode}
                      onChange={(e) => setTransactionCode(e.target.value)}
                      className="w-full p-2 rounded-lg border border-slate-300 bg-white text-sm font-semibold"
                    />
                  </div>
                )}

                <button
                  onClick={handleCompleteSale}
                  disabled={cart.length === 0}
                  className="w-full py-3 bg-green-600 hover:bg-green-700 disabled:bg-slate-300 text-white font-bold rounded-xl shadow transition"
                >
                  Guardar Venta
                </button>

                {saleSuccess && (
                  <div className="flex items-center justify-center gap-2 text-green-800 bg-green-100 p-2.5 rounded-xl text-sm font-semibold">
                    <CheckCircle2 size={18} /> ¡Venta guardada e inventario
                    descontado!
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* 2. INVENTARIO DIARIO Y PRECIOS / STOCK INICIAL PROTEGIDOS */}
        {activeTab === "inventory" && (
          <div className="space-y-6">
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
              <div>
                <h2 className="text-lg font-bold">
                  Inventario Diario y Control de Stock
                </h2>
                <p className="text-sm text-slate-500">
                  Puedes vincular qué insumos ocupa cada producto: al venderlo,
                  se descontarán esos insumos del inventario.
                </p>
              </div>

              <div className="flex flex-wrap gap-2.5">
                <button
                  type="button"
                  onClick={openUpdateInventoryModal}
                  className="flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-xl text-sm font-bold shadow-sm transition"
                >
                  <ClipboardCheck size={17} /> Actualizar inventario
                </button>

                <button
                  type="button"
                  onClick={() =>
                    priceAndStockUnlocked
                      ? setPriceAndStockUnlocked(false)
                      : openPasswordModal("priceAndStock")
                  }
                  className={`flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-sm font-semibold border transition ${
                    priceAndStockUnlocked
                      ? "bg-amber-100 border-amber-400 text-amber-900"
                      : "bg-slate-100 border-slate-300 text-slate-700 hover:bg-slate-200"
                  }`}
                >
                  {priceAndStockUnlocked ? (
                    <Unlock size={16} />
                  ) : (
                    <Lock size={16} />
                  )}
                  {priceAndStockUnlocked
                    ? "Precio / Stock Desbloqueados"
                    : "Desbloquear Precio / Stock"}
                </button>

                <button
                  type="button"
                  onClick={exportToExcel}
                  className="flex items-center gap-1.5 bg-slate-800 hover:bg-slate-900 text-white px-4 py-2 rounded-xl text-sm font-semibold transition"
                >
                  <Download size={16} /> Exportar Excel
                </button>
              </div>
            </div>

            {/* Productos del Menú */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-4">
              <h3 className="font-bold text-slate-800">
                Productos del Menú e Insumos Vinculados (Protegido con
                Negocio2026)
              </h3>

              <form
                onSubmit={handleAddProduct}
                className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-3"
              >
                <div className="flex flex-wrap gap-3 items-end">
                  <div className="flex-1 min-w-[180px]">
                    <label className="block text-xs font-semibold text-slate-600 mb-1">
                      Nombre
                    </label>
                    <input
                      type="text"
                      required
                      value={newProdName}
                      onChange={(e) => setNewProdName(e.target.value)}
                      placeholder="Nombre del producto"
                      className="w-full p-2 border rounded-lg bg-white text-sm"
                    />
                  </div>
                  <div className="w-36">
                    <label className="block text-xs font-semibold text-slate-600 mb-1">
                      Categoría
                    </label>
                    <select
                      value={newProdCategory}
                      onChange={(e) => setNewProdCategory(e.target.value)}
                      className="w-full p-2 border rounded-lg bg-white text-sm"
                    >
                      <option value="Minutas">Minutas</option>
                      <option value="Licuados">Licuados</option>
                    </select>
                  </div>
                  <div className="w-28">
                    <label className="block text-xs font-semibold text-slate-600 mb-1">
                      Precio ($)
                    </label>
                    <input
                      type="number"
                      step="0.01"
                      required
                      value={newProdPrice}
                      onChange={(e) => setNewProdPrice(e.target.value)}
                      placeholder="2.00"
                      className="w-full p-2 border rounded-lg bg-white text-sm"
                    />
                  </div>
                  <div className="w-32">
                    <label className="block text-xs font-semibold text-slate-600 mb-1">
                      Stock Inicial
                    </label>
                    <input
                      type="number"
                      value={newProdStock}
                      onChange={(e) => setNewProdStock(e.target.value)}
                      placeholder={newProdCategory === "Minutas" ? "400" : "600"}
                      className="w-full p-2 border rounded-lg bg-white text-sm"
                    />
                  </div>
                  <div className="flex-1 min-w-[160px]">
                    <label className="block text-xs font-semibold text-slate-600 mb-1">
                      Descripción
                    </label>
                    <input
                      type="text"
                      value={newProdDesc}
                      onChange={(e) => setNewProdDesc(e.target.value)}
                      placeholder="Detalle opcional"
                      className="w-full p-2 border rounded-lg bg-white text-sm"
                    />
                  </div>
                </div>

                <div className="pt-2 border-t border-slate-200 space-y-2">
                  <label className="block text-xs font-bold uppercase text-slate-600">
                    ¿Qué insumos ocupa este producto? (se restarán al vender)
                  </label>
                  <div className="flex flex-wrap gap-2 items-end">
                    <div className="flex-1 min-w-[200px]">
                      <select
                        value={newSelectedSupplyId}
                        onChange={(e) => setNewSelectedSupplyId(e.target.value)}
                        className="w-full p-2 border rounded-lg bg-white text-sm"
                      >
                        {supplies.map((s) => (
                          <option key={s.id} value={s.id}>
                            {s.name} (Disp: {s.stock})
                          </option>
                        ))}
                      </select>
                    </div>
                    <div className="w-28">
                      <input
                        type="number"
                        step="any"
                        min="0.01"
                        value={newSelectedSupplyQty}
                        onChange={(e) => setNewSelectedSupplyQty(e.target.value)}
                        placeholder="Cant."
                        className="w-full p-2 border rounded-lg bg-white text-sm"
                      />
                    </div>
                    <button
                      type="button"
                      onClick={() =>
                        addSupplyToRecipe(
                          newProdRecipe,
                          setNewProdRecipe,
                          newSelectedSupplyId,
                          newSelectedSupplyQty
                        )
                      }
                      className="bg-slate-800 hover:bg-slate-900 text-white px-3 py-2 rounded-lg text-xs font-bold flex items-center gap-1"
                    >
                      <Plus size={14} /> Agregar Insumo
                    </button>
                  </div>

                  {newProdRecipe.length > 0 && (
                    <div className="flex flex-wrap gap-2 pt-1">
                      {newProdRecipe.map((r) => (
                        <span
                          key={r.supplyId}
                          className="inline-flex items-center gap-1.5 bg-blue-50 border border-blue-200 text-blue-900 text-xs font-semibold px-2.5 py-1 rounded-lg"
                        >
                          {r.qty}x {getSupplyName(r.supplyId)}
                          <button
                            type="button"
                            onClick={() =>
                              removeSupplyFromRecipe(
                                newProdRecipe,
                                setNewProdRecipe,
                                r.supplyId
                              )
                            }
                            className="text-red-600 hover:text-red-800"
                          >
                            <X size={13} />
                          </button>
                        </span>
                      ))}
                    </div>
                  )}
                </div>

                <div className="flex justify-end pt-1">
                  <button
                    type="submit"
                    className="bg-blue-700 hover:bg-blue-800 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1"
                  >
                    <Plus size={16} /> Agregar Producto al Menú
                  </button>
                </div>
              </form>

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-sm">
                  <thead className="bg-slate-50 border-b text-xs uppercase text-slate-500">
                    <tr>
                      <th className="p-3">Producto</th>
                      <th className="p-3">Insumos Vinculados</th>
                      <th className="p-3">Precio ($)</th>
                      <th className="p-3">Stock / Disp.</th>
                      <th className="p-3 text-right">Acciones</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y">
                    {products.map((prod) => (
                      <tr key={prod.id} className="hover:bg-slate-50">
                        <td className="p-3">
                          <p className="font-semibold text-slate-900">
                            {prod.name}
                          </p>
                          <p className="text-xs text-slate-500">
                            {prod.category}
                          </p>
                        </td>
                        <td className="p-3">
                          {Array.isArray(prod.recipe) &&
                          prod.recipe.length > 0 ? (
                            <div className="flex flex-wrap gap-1">
                              {prod.recipe.map((r) => (
                                <span
                                  key={r.supplyId}
                                  className="text-xs bg-slate-100 border border-slate-200 px-2 py-0.5 rounded-md font-medium"
                                >
                                  {r.qty}x {getSupplyName(r.supplyId)}
                                </span>
                              ))}
                            </div>
                          ) : (
                            <span className="text-xs text-slate-400">
                              Base por categoría
                            </span>
                          )}
                        </td>
                        <td className="p-3">
                          <div className="flex items-center gap-1.5">
                            <span className="text-slate-400">$</span>
                            <input
                              type="number"
                              step="0.01"
                              readOnly={!priceAndStockUnlocked}
                              onClick={() => {
                                if (!priceAndStockUnlocked)
                                  openPasswordModal("priceAndStock");
                              }}
                              value={prod.price}
                              onChange={(e) =>
                                handleProtectedPriceChange(
                                  prod.id,
                                  e.target.value
                                )
                              }
                              className={`w-24 p-1.5 border rounded-lg font-bold ${
                                priceAndStockUnlocked
                                  ? "bg-white border-amber-400"
                                  : "bg-slate-100 border-slate-200 cursor-pointer"
                              }`}
                            />
                            {!priceAndStockUnlocked && (
                              <Lock size={14} className="text-slate-400" />
                            )}
                          </div>
                        </td>
                        <td className="p-3">
                          <div className="flex items-center gap-1.5">
                            <input
                              type="number"
                              readOnly={!priceAndStockUnlocked}
                              onClick={() => {
                                if (!priceAndStockUnlocked)
                                  openPasswordModal("priceAndStock");
                              }}
                              value={prod.stock}
                              onChange={(e) =>
                                handleProtectedStockChange(
                                  prod.id,
                                  e.target.value
                                )
                              }
                              className={`w-24 p-1.5 border rounded-lg font-bold ${
                                priceAndStockUnlocked
                                  ? "bg-white border-amber-400"
                                  : "bg-slate-100 border-slate-200 cursor-pointer"
                              }`}
                            />
                            {!priceAndStockUnlocked && (
                              <Lock size={14} className="text-slate-400" />
                            )}
                          </div>
                        </td>
                        <td className="p-3 text-right whitespace-nowrap">
                          <div className="inline-flex items-center gap-1">
                            <button
                              type="button"
                              onClick={() => handleRequestEditProduct(prod)}
                              className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-slate-300 text-xs font-bold text-slate-700 hover:bg-slate-100"
                            >
                              {!priceAndStockUnlocked && <Lock size={12} />}
                              <Pencil size={13} /> Editar
                            </button>
                            <button
                              type="button"
                              onClick={() =>
                                setProducts(
                                  products.filter((p) => p.id !== prod.id)
                                )
                              }
                              className="text-red-600 hover:bg-red-50 p-1.5 rounded-lg"
                            >
                              <Trash2 size={16} />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Materiales e Insumos (Crear, Listar y Eliminar) */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-4">
              <div className="flex flex-wrap justify-between items-center gap-2">
                <div>
                  <h3 className="font-bold text-slate-800">
                    Materiales e Insumos (Empaques y Materia Prima)
                  </h3>
                  <p className="text-xs text-slate-500">
                    Crea o elimina insumos, ajusta cantidades con{" "}
                    <strong>Actualizar inventario</strong> o descuéntalos al
                    vender.
                  </p>
                </div>
                <span className="inline-flex items-center gap-1 text-xs font-semibold bg-slate-100 text-slate-700 px-3 py-1 rounded-full">
                  <Lock size={13} /> Protegido con Negocio2026
                </span>
              </div>

              <form
                onSubmit={handleAddSupply}
                className="flex flex-wrap gap-3 items-end bg-slate-50 p-3.5 rounded-xl border border-slate-200"
              >
                <div className="flex-1 min-w-[180px]">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">
                    Nombre del insumo
                  </label>
                  <input
                    type="text"
                    required
                    value={newSupplyName}
                    onChange={(e) => setNewSupplyName(e.target.value)}
                    placeholder="Ej. Leche entera, Vasos 12oz..."
                    className="w-full p-2 border rounded-lg bg-white text-sm"
                  />
                </div>
                <div className="w-40">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">
                    Tipo
                  </label>
                  <select
                    value={newSupplyType}
                    onChange={(e) => setNewSupplyType(e.target.value)}
                    className="w-full p-2 border rounded-lg bg-white text-sm"
                  >
                    <option value="Materia Prima">Materia Prima</option>
                    <option value="Empaques">Empaques</option>
                  </select>
                </div>
                <div className="w-32">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">
                    Stock inicial
                  </label>
                  <input
                    type="number"
                    step="any"
                    min="0"
                    value={newSupplyStock}
                    onChange={(e) => setNewSupplyStock(e.target.value)}
                    placeholder="0"
                    className="w-full p-2 border rounded-lg bg-white text-sm"
                  />
                </div>
                <button
                  type="submit"
                  className="bg-blue-700 hover:bg-blue-800 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1"
                >
                  {!priceAndStockUnlocked && <Lock size={14} />}
                  <Plus size={16} /> Crear Insumo
                </button>
              </form>

              <div className="overflow-x-auto max-h-96">
                <table className="w-full text-left border-collapse text-sm">
                  <thead className="bg-slate-50 border-b text-xs uppercase text-slate-500 sticky top-0">
                    <tr>
                      <th className="p-3">Detalle</th>
                      <th className="p-3">Tipo</th>
                      <th className="p-3">Total disponible</th>
                      <th className="p-3 text-right">Eliminar</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y">
                    {supplies.map((sup) => (
                      <tr key={sup.id} className="hover:bg-slate-50">
                        <td className="p-3 font-medium">{sup.name}</td>
                        <td className="p-3 text-slate-600">{sup.type}</td>
                        <td className="p-3 font-extrabold text-slate-900">
                          {sup.stock}
                        </td>
                        <td className="p-3 text-right">
                          <button
                            type="button"
                            onClick={() => handleDeleteSupply(sup)}
                            className="inline-flex items-center gap-1 text-red-600 hover:bg-red-50 px-2 py-1 rounded-lg text-xs font-bold"
                          >
                            {!priceAndStockUnlocked && <Lock size={12} />}
                            <Trash2 size={15} />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Historial de Actualizaciones de Inventario */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-4">
              <h3 className="font-bold text-slate-800">
                Historial de Actualizaciones de Inventario
              </h3>

              {inventoryLogs.length === 0 ? (
                <p className="text-sm text-slate-400 py-4">
                  Aún no hay actualizaciones registradas. Usa el botón{" "}
                  <strong>Actualizar inventario</strong> para registrar una.
                </p>
              ) : (
                <div className="space-y-3">
                  {inventoryLogs.map((log) => (
                    <div
                      key={log.id}
                      className="p-4 rounded-xl border border-slate-200 bg-slate-50 space-y-2"
                    >
                      <div className="flex flex-wrap justify-between items-center gap-2 text-xs">
                        <div className="flex items-center gap-2 font-bold text-slate-800">
                          <Calendar size={14} className="text-blue-700" />
                          <span>Fecha: {log.date}</span>
                          <span className="px-2 py-0.5 rounded-full bg-blue-100 text-blue-800 uppercase">
                            {log.mode}
                          </span>
                        </div>
                        <span className="text-slate-500">
                          Registrado: {log.recordedAt}
                        </span>
                      </div>

                      {log.note && (
                        <p className="text-xs text-slate-600 italic">
                          Nota: {log.note}
                        </p>
                      )}

                      <div className="flex flex-wrap gap-2 pt-1">
                        {log.items.map((it, idx) => (
                          <span
                            key={`${log.id}-${idx}`}
                            className="text-xs bg-white border border-slate-200 px-2.5 py-1 rounded-lg font-medium"
                          >
                            <strong>{it.name}:</strong> {it.previousStock} →{" "}
                            <span className="text-blue-800 font-bold">
                              {it.newStock}
                            </span>{" "}
                            ({it.delta >= 0 ? `+${it.delta}` : it.delta})
                          </span>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {/* 3. HISTORIAL DE VENTAS CON MODIFICACIÓN Y ELIMINACIÓN PROTEGIDAS */}
        {activeTab === "history" && (
          <div className="space-y-6">
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
              <div>
                <h2 className="text-lg font-bold">Historial de Ventas</h2>
                <p className="text-2xl font-extrabold text-green-700 mt-1">
                  Total Vendido: ${totalSold.toFixed(2)}
                </p>
              </div>
              <div className="flex flex-wrap gap-2.5">
                <button
                  type="button"
                  onClick={() =>
                    salesEditUnlocked
                      ? setSalesEditUnlocked(false)
                      : openPasswordModal("editSale")
                  }
                  className={`flex items-center gap-1.5 px-3.5 py-2.5 rounded-xl text-sm font-semibold border transition ${
                    salesEditUnlocked
                      ? "bg-amber-100 border-amber-400 text-amber-900"
                      : "bg-slate-100 border-slate-300 text-slate-700 hover:bg-slate-200"
                  }`}
                >
                  {salesEditUnlocked ? (
                    <Unlock size={16} />
                  ) : (
                    <Lock size={16} />
                  )}
                  {salesEditUnlocked
                    ? "Edición Desbloqueada"
                    : "Desbloquear Edición"}
                </button>

                <button
                  onClick={exportToExcel}
                  className="flex items-center gap-2 bg-slate-800 hover:bg-slate-900 text-white px-4 py-2.5 rounded-xl text-sm font-semibold"
                >
                  <Download size={17} /> Exportar Reporte Excel
                </button>
              </div>
            </div>

            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-x-auto">
              {sales.length === 0 ? (
                <p className="text-center text-slate-400 py-10">
                  No hay ventas registradas todavía.
                </p>
              ) : (
                <table className="w-full text-left border-collapse text-sm">
                  <thead className="bg-slate-50 border-b text-xs uppercase text-slate-500">
                    <tr>
                      <th className="p-4">Fecha y Hora</th>
                      <th className="p-4">Productos</th>
                      <th className="p-4">Pago</th>
                      <th className="p-4">Total</th>
                      <th className="p-4">Cambio</th>
                      <th className="p-4 text-right">Acciones</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y">
                    {sales.map((sale) => (
                      <tr key={sale.id} className="hover:bg-slate-50">
                        <td className="p-4 text-slate-500 whitespace-nowrap">
                          {sale.date}
                          {sale.editedAt && (
                            <span className="block text-[11px] text-amber-700">
                              Editada: {sale.editedAt}
                            </span>
                          )}
                        </td>
                        <td className="p-4 font-medium">
                          {sale.items
                            .map((i) => `${i.qty}x ${i.name}`)
                            .join(", ")}
                        </td>
                        <td className="p-4">
                          <span className="px-2.5 py-1 rounded-full text-xs font-semibold bg-blue-100 text-blue-800">
                            {sale.paymentMethod}
                          </span>
                          {(sale.paymentMethod === "Transferencia" ||
                            sale.paymentMethod === "Tarjeta") &&
                            sale.transactionCode && (
                              <span className="block text-[11px] text-slate-600 font-semibold mt-1">
                                Cód: {sale.transactionCode}
                              </span>
                            )}
                        </td>
                        <td className="p-4 font-extrabold text-slate-900">
                          ${sale.total.toFixed(2)}
                        </td>
                        <td className="p-4 text-slate-600">
                          {sale.paymentMethod === "Efectivo"
                            ? `$${sale.change.toFixed(2)}`
                            : "-"}
                        </td>
                        <td className="p-4 text-right whitespace-nowrap">
                          <div className="inline-flex items-center gap-1.5">
                            <button
                              type="button"
                              onClick={() => handleRequestEditSale(sale)}
                              className="inline-flex items-center gap-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs font-bold text-slate-700 hover:bg-slate-100"
                            >
                              {!salesEditUnlocked && <Lock size={13} />}
                              <Pencil size={13} /> Modificar
                            </button>
                            <button
                              type="button"
                              onClick={() => handleRequestDeleteSale(sale)}
                              className="inline-flex items-center gap-1 px-3 py-1.5 rounded-lg border border-red-200 bg-red-50 text-xs font-bold text-red-700 hover:bg-red-100"
                            >
                              {!salesEditUnlocked && <Lock size={13} />}
                              <Trash2 size={13} /> Eliminar
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        )}

        {/* 4. FINANCIERO (PROTEGIDO CON CONTRASEÑA) */}
        {activeTab === "financial" && financialUnlocked && (
          <div className="space-y-6">
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
              <div className="flex items-center gap-2.5">
                <Wallet className="text-blue-800" size={24} />
                <div>
                  <h2 className="text-lg font-bold">Control Financiero</h2>
                  <p className="text-xs text-slate-500">
                    Meta general, efectivo neto descontando reinversión, fechas
                    de Kat y Gab, deudas y pagos
                  </p>
                </div>
              </div>

              <button
                onClick={() => {
                  setFinancialUnlocked(false);
                  setActiveTab("pos");
                }}
                className="flex items-center gap-1.5 px-3.5 py-2 rounded-xl border border-slate-300 text-sm font-semibold text-slate-700 hover:bg-slate-50"
              >
                <EyeOff size={16} /> Bloquear
              </button>
            </div>

            {/* BLOQUE DEDICADO A LA META GENERAL ($7,000) */}
            <div className="bg-gradient-to-r from-blue-900 to-blue-800 text-white p-6 rounded-2xl shadow-md space-y-4">
              <div className="flex flex-wrap justify-between items-start gap-4">
                <div className="flex items-center gap-2.5">
                  <Target className="text-amber-300" size={26} />
                  <div>
                    <h3 className="text-lg font-bold">
                      Meta General del Negocio
                    </h3>
                    <p className="text-xs text-blue-200">
                      Resultado neto tras restar deudas, reinversiones y pagos
                      en general
                    </p>
                  </div>
                </div>

                <div className="bg-blue-950/60 px-3.5 py-2 rounded-xl border border-blue-700">
                  <label className="block text-[11px] uppercase text-blue-200 font-semibold">
                    Objetivo Meta ($)
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    value={financial.goal}
                    onChange={(e) =>
                      setFinancial({
                        ...financial,
                        goal: parseFloat(e.target.value) || 0,
                      })
                    }
                    className="bg-transparent text-lg font-extrabold text-amber-300 w-32 focus:outline-none"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3 text-sm">
                <div className="bg-white/10 p-3 rounded-xl">
                  <p className="text-xs text-blue-200">Total Vendido</p>
                  <p className="text-lg font-bold mt-0.5">
                    +${totalSold.toFixed(2)}
                  </p>
                </div>

                <div className="bg-white/10 p-3 rounded-xl">
                  <p className="text-xs text-blue-200">(-) Deudas Pendientes</p>
                  <p className="text-lg font-bold text-red-300 mt-0.5">
                    -${totalPendingDebtsAndSalaries.toFixed(2)}
                  </p>
                </div>

                <div className="bg-white/10 p-3 rounded-xl">
                  <p className="text-xs text-blue-200">(-) Reinversión</p>
                  <p className="text-lg font-bold text-amber-200 mt-0.5">
                    -${totalReinvestments.toFixed(2)}
                  </p>
                </div>

                <div className="bg-white/10 p-3 rounded-xl">
                  <p className="text-xs text-blue-200">(-) Pagos Realizados</p>
                  <p className="text-lg font-bold text-orange-200 mt-0.5">
                    -${totalPaymentsMade.toFixed(2)}
                  </p>
                </div>

                <div className="bg-white text-slate-900 p-3 rounded-xl shadow">
                  <p className="text-xs font-bold uppercase text-slate-500">
                    Resultado Neto Meta
                  </p>
                  <p
                    className={`text-xl font-extrabold mt-0.5 ${
                      netTowardGoal >= 0 ? "text-emerald-700" : "text-red-600"
                    }`}
                  >
                    ${netTowardGoal.toFixed(2)}
                  </p>
                  <p className="text-[11px] text-slate-500">
                    Falta para meta: ${remainingForGoal.toFixed(2)}
                  </p>
                </div>
              </div>

              <div className="space-y-1">
                <div className="flex justify-between text-xs text-blue-100">
                  <span>Progreso hacia ${goalTarget.toFixed(2)}</span>
                  <span>{goalProgressPct.toFixed(1)}%</span>
                </div>
                <div className="w-full h-2.5 bg-blue-950 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-amber-400 transition-all"
                    style={{ width: `${goalProgressPct}%` }}
                  />
                </div>
              </div>
            </div>

            {/* Indicadores de Efectivo y Reinversión */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  Ingresos en Efectivo
                </p>
                <p className="mt-1 text-2xl font-extrabold text-slate-900">
                  ${totalCashSales.toFixed(2)}
                </p>
                <p className="text-[11px] text-slate-500 mt-1">
                  Transferencia/Tarjeta: ${totalDigitalSales.toFixed(2)}
                </p>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  (-) Reinversión Restada de Efectivo
                </p>
                <p className="mt-1 text-2xl font-extrabold text-amber-700">
                  -${totalReinvestments.toFixed(2)}
                </p>
                <p className="text-[11px] text-slate-500 mt-1">
                  Se descuenta directo del efectivo ingresado
                </p>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  Efectivo Neto en Caja
                </p>
                <p
                  className={`mt-1 text-2xl font-extrabold ${
                    netCashInDrawer >= 0 ? "text-emerald-700" : "text-red-600"
                  }`}
                >
                  ${netCashInDrawer.toFixed(2)}
                </p>
                <p className="text-[11px] text-slate-500 mt-1">
                  Efectivo ingresado − reinversión − pagos
                </p>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  Saldo Total Disponible
                </p>
                <p
                  className={`mt-1 text-2xl font-extrabold ${
                    availableCash > 0 ? "text-blue-800" : "text-slate-600"
                  }`}
                >
                  ${availableCash.toFixed(2)}
                </p>
                <p className="text-[11px] text-slate-500 mt-1">
                  {availableCash > 0
                    ? "Hay dinero disponible para cubrir pagos"
                    : "Sin saldo positivo disponible aún"}
                </p>
              </div>
            </div>

            {/* Fechas y Pagos de Kat y Gab */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-4">
              <div>
                <h3 className="text-base font-bold text-slate-900">
                  Pagos y Fechas — Kat y Gab
                </h3>
                <p className="text-xs text-slate-500">
                  Inicio: <strong>Kat el 3 de diciembre</strong> y{" "}
                  <strong>Gab el 4 de diciembre</strong>. Fechas de pago:{" "}
                  <strong>15 de diciembre</strong> y{" "}
                  <strong>30 de diciembre</strong>.
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {financial.employees.map((emp) => (
                  <div
                    key={emp.id}
                    className="p-4 rounded-2xl border border-slate-200 bg-slate-50 space-y-3"
                  >
                    <div className="flex justify-between items-center">
                      <h4 className="font-bold text-slate-900 text-base">
                        {emp.name}
                      </h4>
                      <span className="text-xs font-semibold bg-blue-100 text-blue-800 px-2.5 py-0.5 rounded-full">
                        Inicia: {emp.startDate}
                      </span>
                    </div>

                    <div className="space-y-2">
                      {emp.payDates.map((pd) => {
                        const canCoverNow =
                          !pd.paid && availableCash >= pd.amount;
                        return (
                          <div
                            key={pd.id}
                            className={`p-3 rounded-xl border flex flex-wrap justify-between items-center gap-2 ${
                              pd.paid
                                ? "bg-green-50 border-green-200"
                                : canCoverNow
                                ? "bg-emerald-50 border-emerald-300"
                                : "bg-white border-slate-200"
                            }`}
                          >
                            <div>
                              <p className="text-sm font-bold text-slate-800">
                                {pd.label} — Fecha: {pd.dueDate}
                              </p>
                              <p className="text-xs text-slate-500">
                                Monto: <strong>${pd.amount.toFixed(2)}</strong>
                                {pd.note ? ` (${pd.note})` : ""}
                                {pd.paid && pd.paidDate
                                  ? ` · Pagado el ${pd.paidDate}`
                                  : ""}
                              </p>
                              {!pd.paid && canCoverNow && (
                                <p className="text-[11px] font-semibold text-emerald-700 mt-0.5">
                                  ✓ Dinero disponible para pagar este monto
                                </p>
                              )}
                            </div>

                            <button
                              type="button"
                              onClick={() =>
                                toggleEmployeePayment(emp.id, pd.id)
                              }
                              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                                pd.paid
                                  ? "bg-slate-200 text-slate-700 hover:bg-slate-300"
                                  : canCoverNow
                                  ? "bg-green-600 text-white hover:bg-green-700"
                                  : "bg-blue-700 text-white hover:bg-blue-800"
                              }`}
                            >
                              {pd.paid ? "Marcar Pendiente" : "Registrar Pago"}
                            </button>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </div>

              <form
                onSubmit={handleAddEmployeePayDate}
                className="flex flex-wrap gap-3 items-end bg-slate-50 p-3.5 rounded-xl border border-slate-200"
              >
                <div className="w-36">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">
                    Empleada
                  </label>
                  <select
                    value={extraPayEmpId}
                    onChange={(e) => setExtraPayEmpId(e.target.value)}
                    className="w-full p-2 border rounded-lg bg-white text-sm"
                  >
                    <option value="emp-kat">Kat</option>
                    <option value="emp-gab">Gab</option>
                  </select>
                </div>

                <div className="flex-1 min-w-[160px]">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">
                    Concepto de pago
                  </label>
                  <input
                    type="text"
                    value={extraPayLabel}
                    onChange={(e) => setExtraPayLabel(e.target.value)}
                    className="w-full p-2 border rounded-lg bg-white text-sm"
                  />
                </div>

                <div className="w-40">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">
                    Fecha de pago
                  </label>
                  <input
                    type="date"
                    required
                    value={extraPayDate}
                    onChange={(e) => setExtraPayDate(e.target.value)}
                    className="w-full p-2 border rounded-lg bg-white text-sm"
                  />
                </div>

                <div className="w-32">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">
                    Monto ($)
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    required
                    value={extraPayAmount}
                    onChange={(e) => setExtraPayAmount(e.target.value)}
                    placeholder="0.00"
                    className="w-full p-2 border rounded-lg bg-white text-sm"
                  />
                </div>

                <button
                  type="submit"
                  className="bg-blue-700 hover:bg-blue-800 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1"
                >
                  <Plus size={16} /> Agregar Fecha
                </button>
              </form>
            </div>

            {/* Deudas y Reinversiones */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-4">
                <h3 className="font-bold text-slate-900">
                  Deudas y Compromisos
                </h3>

                <form
                  onSubmit={handleAddDebt}
                  className="flex flex-wrap gap-2 items-end bg-slate-50 p-3 rounded-xl border border-slate-200"
                >
                  <input
                    type="text"
                    required
                    placeholder="Concepto de deuda"
                    value={newDebtConcept}
                    onChange={(e) => setNewDebtConcept(e.target.value)}
                    className="flex-1 min-w-[140px] p-2 border rounded-lg bg-white text-sm"
                  />
                  <input
                    type="date"
                    value={newDebtDate}
                    onChange={(e) => setNewDebtDate(e.target.value)}
                    className="p-2 border rounded-lg bg-white text-sm"
                  />
                  <input
                    type="number"
                    step="0.01"
                    required
                    placeholder="Monto ($)"
                    value={newDebtAmount}
                    onChange={(e) => setNewDebtAmount(e.target.value)}
                    className="w-28 p-2 border rounded-lg bg-white text-sm"
                  />
                  <button
                    type="submit"
                    className="bg-red-600 hover:bg-red-700 text-white px-3.5 py-2 rounded-lg text-sm font-semibold"
                  >
                    Registrar
                  </button>
                </form>

                <div className="space-y-2 max-h-72 overflow-y-auto">
                  {financial.otherDebts.map((debt) => {
                    const canPayDebt =
                      !debt.paid &&
                      availableCash >= debt.amount &&
                      debt.amount > 0;
                    return (
                      <div
                        key={debt.id}
                        className="p-3 rounded-xl border border-slate-200 flex justify-between items-center gap-2 text-sm"
                      >
                        <div>
                          <p className="font-bold text-slate-800">
                            {debt.concept}
                          </p>
                          <p className="text-xs text-slate-500">
                            Fecha: {debt.date || "Sin fecha"} ·{" "}
                            <strong className="text-slate-800">
                              ${Number(debt.amount).toFixed(2)}
                            </strong>
                          </p>
                          {canPayDebt && (
                            <p className="text-[11px] font-semibold text-emerald-700 flex items-center gap-1 mt-0.5">
                              <AlertCircle size={12} /> Hay fondos disponibles
                              para pagar
                            </p>
                          )}
                        </div>

                        <button
                          type="button"
                          onClick={() => toggleDebtPaid(debt.id)}
                          className={`px-3 py-1 rounded-lg text-xs font-bold ${
                            debt.paid
                              ? "bg-green-100 text-green-800"
                              : "bg-slate-800 text-white hover:bg-slate-900"
                          }`}
                        >
                          {debt.paid ? "Pagada" : "Marcar Pagada"}
                        </button>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-4">
                <div>
                  <h3 className="font-bold text-slate-900">
                    Registro de Reinversiones
                  </h3>
                  <p className="text-xs text-slate-500">
                    Cada reinversión se resta automáticamente del efectivo
                    ingresado y de la meta general.
                  </p>
                </div>

                <form
                  onSubmit={handleAddReinvestment}
                  className="flex flex-wrap gap-2 items-end bg-slate-50 p-3 rounded-xl border border-slate-200"
                >
                  <input
                    type="text"
                    required
                    placeholder="Detalle de reinversión"
                    value={newReinvConcept}
                    onChange={(e) => setNewReinvConcept(e.target.value)}
                    className="flex-1 min-w-[140px] p-2 border rounded-lg bg-white text-sm"
                  />
                  <select
                    value={newReinvCategory}
                    onChange={(e) => setNewReinvCategory(e.target.value)}
                    className="p-2 border rounded-lg bg-white text-sm"
                  >
                    <option value="Minutas">Minutas</option>
                    <option value="Licuados">Licuados</option>
                    <option value="Bebidas">Bebidas</option>
                    <option value="General">General</option>
                  </select>
                  <input
                    type="date"
                    value={newReinvDate}
                    onChange={(e) => setNewReinvDate(e.target.value)}
                    className="p-2 border rounded-lg bg-white text-sm"
                  />
                  <input
                    type="number"
                    step="0.01"
                    required
                    placeholder="Monto ($)"
                    value={newReinvAmount}
                    onChange={(e) => setNewReinvAmount(e.target.value)}
                    className="w-28 p-2 border rounded-lg bg-white text-sm"
                  />
                  <button
                    type="submit"
                    className="bg-blue-700 hover:bg-blue-800 text-white px-3.5 py-2 rounded-lg text-sm font-semibold"
                  >
                    Guardar
                  </button>
                </form>

                <div className="space-y-2 max-h-72 overflow-y-auto">
                  {financial.reinvestments.length === 0 ? (
                    <p className="text-xs text-slate-400 py-6 text-center">
                      No hay reinversiones registradas todavía.
                    </p>
                  ) : (
                    financial.reinvestments.map((reinv) => (
                      <div
                        key={reinv.id}
                        className="p-3 rounded-xl border border-slate-200 flex justify-between items-center text-sm"
                      >
                        <div>
                          <p className="font-bold text-slate-800">
                            {reinv.concept}
                          </p>
                          <p className="text-xs text-slate-500">
                            {reinv.category} · Fecha: {reinv.date}
                          </p>
                        </div>
                        <span className="font-extrabold text-amber-700">
                          -${Number(reinv.amount).toFixed(2)}
                        </span>
                      </div>
                    ))
                  )}
                </div>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}