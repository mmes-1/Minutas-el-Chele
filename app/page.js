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
  Upload,
  Download,
  Plus,
  Minus,
  Trash2,
  DollarSign,
  CreditCard,
  ArrowRightLeft,
  Search,
  CheckCircle2,
  Eye,
  EyeOff,
} from "lucide-react";

const OWNER_PASSWORD = "Negocio2026";

const INITIAL_MENU = [
  {
    id: 1,
    name: "Ceviminuta",
    category: "Minutas",
    price: 5.0,
    stock: 40,
    description:
      "Minuta de limon, sal, chile, chamoy, salsa negra, tajin, alguashte, nachos, mango o pepino, chimol, camarones",
  },
  {
    id: 2,
    name: "Limon y fruta",
    category: "Minutas",
    price: 3.0,
    stock: 50,
    description:
      "Minuta de limon, sal, chile, chamoy, salsa negra, tajin, alguashte, nachos, mango o pepino",
  },
  {
    id: 3,
    name: "Tadicional con gomitas",
    category: "Minutas",
    price: 3.0,
    stock: 50,
    description:
      "Minuta de Jarabes de tu elección, leche condensada, gomitas y pixie",
  },
  {
    id: 4,
    name: "Tadicional con fruta",
    category: "Minutas",
    price: 2.0,
    stock: 50,
    description:
      "Minuta de Jarabes de tu elección con fruta picada (guineo y piña) y jalea de tamarindo o leche condensada",
  },
  {
    id: 5,
    name: "Tradicional de limon",
    category: "Minutas",
    price: 1.5,
    stock: 60,
    description:
      "Minuta de limon, sal, chile, chamoy, salsa negra, tajin, alguashte",
  },
  {
    id: 6,
    name: "Licuado de guineo",
    category: "Licuados",
    price: 2.0,
    stock: 40,
    description: "Licuado de guineo preparado con leche",
  },
  {
    id: 7,
    name: "Licuado de fresa",
    category: "Licuados",
    price: 2.0,
    stock: 40,
    description: "Licuado de fresa preparado con leche",
  },
  {
    id: 8,
    name: "Licuado de Oreo",
    category: "Licuados",
    price: 2.0,
    stock: 40,
    description: "Licuado de galleta Oreo preparado con leche",
  },
];

const INITIAL_SUPPLIES = [
  { id: "s1", name: "vasos de licuados", type: "Empaques", stock: 600 },
  { id: "s2", name: "vasos de minutas", type: "Empaques", stock: 300 },
  { id: "s3", name: "bandeja 11,1x11,1,x6,3", type: "Empaques", stock: 400 },
  { id: "s4", name: "tenedores", type: "Empaques", stock: 300 },
  { id: "s5", name: "cucharas", type: "Empaques", stock: 300 },
  { id: "s6", name: "pajillas sin emboltorio azul", type: "Empaques", stock: 300 },
  { id: "s7", name: "pajillas con emboltorio", type: "Empaques", stock: 340 },
  { id: "s8", name: "servilletas", type: "Empaques", stock: 1000 },
  {
    id: "s9",
    name: "Bolsas para guardar la fruta picada",
    type: "Empaques",
    stock: 300,
  },
  { id: "s10", name: "Jarabes", type: "Materia Prima", stock: 8 },
  { id: "s11", name: "Jalea de tamarindo", type: "Materia Prima", stock: 0 },
  { id: "s12", name: "Guineo", type: "Materia Prima", stock: 36 },
  { id: "s13", name: "Piña", type: "Materia Prima", stock: 3 },
  { id: "s14", name: "Limon", type: "Materia Prima", stock: 200 },
  { id: "s15", name: "Chile", type: "Materia Prima", stock: 1 },
  { id: "s16", name: "Aiuashte", type: "Materia Prima", stock: 1 },
  { id: "s17", name: "Salsa negra", type: "Materia Prima", stock: 1 },
  { id: "s18", name: "Nachos", type: "Materia Prima", stock: 72 },
  { id: "s19", name: "Hielo", type: "Materia Prima", stock: 400 },
  { id: "s20", name: "Pepino", type: "Materia Prima", stock: 25 },
  { id: "s21", name: "Tomate", type: "Materia Prima", stock: 75 },
  { id: "s22", name: "Camaron", type: "Materia Prima", stock: 2000 },
  {
    id: "s23",
    name: "agua alpina en botella 699ml",
    type: "Materia Prima",
    stock: 72,
  },
  { id: "s24", name: "lata sutida soda", type: "Materia Prima", stock: 24 },
  { id: "s25", name: "lata coca cola soda", type: "Materia Prima", stock: 48 },
];

const INITIAL_FINANCIAL = {
  goal: 7000.0,
  reinvestment: 0.0,
  purchasesSummary: [
    {
      id: "p1",
      type: "Financiado",
      utensils: 480.25,
      packaging: 162.45,
      rawMaterial: 345.94,
      transport: 40.0,
      paid: 902.44,
    },
    {
      id: "p2",
      type: "Plazo/tarjeta",
      utensils: 696.9,
      packaging: 0,
      rawMaterial: 0,
      transport: 0,
      paid: 0,
    },
    {
      id: "p3",
      type: "Contado Ever",
      utensils: 100.0,
      packaging: 0,
      rawMaterial: 0,
      transport: 0,
      paid: 100.0,
    },
  ],
  debts: [
    { id: "d1", concept: "Salario Kat. 1", amount: 201.6, note: "13,44 x día" },
    { id: "d2", concept: "Salario Gab. 1", amount: 275.0, note: "18,33 x día" },
    { id: "d3", concept: "Salario Kat. 2", amount: 201.6, note: "13,44 x día" },
    { id: "d4", concept: "Salario Gab. 2", amount: 275.0, note: "18,33 x día" },
    { id: "d5", concept: "Prestamo", amount: 1000.0, note: "" },
    { id: "d6", concept: "Intereses", amount: 200.0, note: "" },
    { id: "d7", concept: "Plazo/tarjeta", amount: 696.9, note: "" },
    { id: "d8", concept: "Contado Ever", amount: 0.0, note: "" },
  ],
};

export default function NegocioElCheleApp() {
  const [activeTab, setActiveTab] = useState("pos");
  const [products, setProducts] = useState([]);
  const [supplies, setSupplies] = useState([]);
  const [sales, setSales] = useState([]);
  const [financial, setFinancial] = useState(INITIAL_FINANCIAL);
  const [search, setSearch] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("Todos");

  // Carrito y cobro
  const [cart, setCart] = useState([]);
  const [paymentMethod, setPaymentMethod] = useState("Efectivo");
  const [cashGiven, setCashGiven] = useState("");
  const [saleSuccess, setSaleSuccess] = useState(false);

  // Seguridad con contraseña para precios y sección Financiero
  const [pricesUnlocked, setPricesUnlocked] = useState(false);
  const [financialUnlocked, setFinancialUnlocked] = useState(false);
  const [passwordInput, setPasswordInput] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [pendingAction, setPendingAction] = useState(null); // 'prices' | 'financial'

  // Formularios de inventario
  const [newProdName, setNewProdName] = useState("");
  const [newProdCategory, setNewProdCategory] = useState("Minutas");
  const [newProdPrice, setNewProdPrice] = useState("");
  const [newProdStock, setNewProdStock] = useState("");

  const [newSupplyName, setNewSupplyName] = useState("");
  const [newSupplyType, setNewSupplyType] = useState("Materia Prima");
  const [newSupplyStock, setNewSupplyStock] = useState("");

  useEffect(() => {
    const savedProducts = localStorage.getItem("chele_v2_products");
    const savedSupplies = localStorage.getItem("chele_v2_supplies");
    const savedSales = localStorage.getItem("chele_v2_sales");
    const savedFinancial = localStorage.getItem("chele_v2_financial");

    setProducts(savedProducts ? JSON.parse(savedProducts) : INITIAL_MENU);
    setSupplies(savedSupplies ? JSON.parse(savedSupplies) : INITIAL_SUPPLIES);
    setSales(savedSales ? JSON.parse(savedSales) : []);
    setFinancial(
      savedFinancial ? JSON.parse(savedFinancial) : INITIAL_FINANCIAL
    );
  }, []);

  useEffect(() => {
    if (products.length > 0) {
      localStorage.setItem("chele_v2_products", JSON.stringify(products));
    }
  }, [products]);

  useEffect(() => {
    if (supplies.length > 0) {
      localStorage.setItem("chele_v2_supplies", JSON.stringify(supplies));
    }
  }, [supplies]);

  useEffect(() => {
    localStorage.setItem("chele_v2_sales", JSON.stringify(sales));
  }, [sales]);

  useEffect(() => {
    localStorage.setItem("chele_v2_financial", JSON.stringify(financial));
  }, [financial]);

  // --- VALIDACIÓN DE CONTRASEÑA ---
  const openPasswordModal = (target) => {
    setPasswordInput("");
    setPasswordError("");
    setPendingAction(target);
  };

  const handlePasswordSubmit = (e) => {
    e.preventDefault();
    if (passwordInput === OWNER_PASSWORD) {
      if (pendingAction === "prices") {
        setPricesUnlocked(true);
      } else if (pendingAction === "financial") {
        setFinancialUnlocked(true);
        setActiveTab("financial");
      }
      setPendingAction(null);
      setPasswordInput("");
      setPasswordError("");
    } else {
      setPasswordError("Contraseña incorrecta. Solo el dueño con Negocio2026 puede acceder.");
    }
  };

  const handleSelectFinancialTab = () => {
    if (financialUnlocked) {
      setActiveTab("financial");
    } else {
      openPasswordModal("financial");
    }
  };

  // --- CARRITO Y VENTAS ---
  const addToCart = (product) => {
    if (product.stock <= 0) return;
    const existing = cart.find((item) => item.id === product.id);
    if (existing) {
      if (existing.qty < product.stock) {
        setCart(
          cart.map((item) =>
            item.id === product.id ? { ...item, qty: item.qty + 1 } : item
          )
        );
      }
    } else {
      setCart([...cart, { ...product, qty: 1 }]);
    }
  };

  const updateCartQty = (id, delta) => {
    const prod = products.find((p) => p.id === id);
    setCart(
      cart
        .map((item) => {
          if (item.id === id) {
            const newQty = item.qty + delta;
            if (newQty <= 0) return null;
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
      alert("Ingresa un monto en efectivo igual o mayor al total a pagar.");
      return;
    }

    const updatedProducts = products.map((prod) => {
      const sold = cart.find((c) => c.id === prod.id);
      if (sold) {
        return { ...prod, stock: Math.max(0, prod.stock - sold.qty) };
      }
      return prod;
    });
    setProducts(updatedProducts);

    const newSale = {
      id: Date.now(),
      date: new Date().toLocaleString("es-SV"),
      items: cart,
      total: cartTotal,
      paymentMethod,
      cashGiven: paymentMethod === "Efectivo" ? cashNum : cartTotal,
      change: paymentMethod === "Efectivo" ? Math.max(0, changeAmount) : 0,
    };

    setSales([newSale, ...sales]);
    setCart([]);
    setCashGiven("");
    setSaleSuccess(true);
    setTimeout(() => setSaleSuccess(false), 3000);
  };

  // --- EDICIÓN PROTEGIDA DE PRECIOS ---
  const handlePriceChange = (productId, rawValue) => {
    if (!pricesUnlocked) {
      openPasswordModal("prices");
      return;
    }
    const parsed = parseFloat(rawValue);
    setProducts(
      products.map((p) =>
        p.id === productId ? { ...p, price: Number.isNaN(parsed) ? 0 : parsed } : p
      )
    );
  };

  const handleProductStockChange = (productId, rawValue) => {
    const parsed = parseInt(rawValue, 10);
    setProducts(
      products.map((p) =>
        p.id === productId ? { ...p, stock: Number.isNaN(parsed) ? 0 : parsed } : p
      )
    );
  };

  const handleAddProduct = (e) => {
    e.preventDefault();
    if (!newProdName.trim()) return;
    const priceVal = parseFloat(newProdPrice);
    if (!pricesUnlocked && priceVal > 0) {
      openPasswordModal("prices");
      return;
    }
    const stockVal = parseInt(newProdStock, 10) || 0;
    setProducts([
      ...products,
      {
        id: Date.now(),
        name: newProdName.trim(),
        category: newProdCategory,
        price: Number.isNaN(priceVal) ? 0 : priceVal,
        stock: stockVal,
        description: "Producto agregado desde inventario",
      },
    ]);
    setNewProdName("");
    setNewProdPrice("");
    setNewProdStock("");
  };

  const handleAddSupply = (e) => {
    e.preventDefault();
    if (!newSupplyName.trim()) return;
    setSupplies([
      ...supplies,
      {
        id: `s-${Date.now()}`,
        name: newSupplyName.trim(),
        type: newSupplyType,
        stock: parseInt(newSupplyStock, 10) || 0,
      },
    ]);
    setNewSupplyName("");
    setNewSupplyStock("");
  };

  // --- IMPORTAR / EXPORTAR EXCEL ---
  const handleExcelUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: "binary" });

      // Si el Excel trae hoja "Menu"
      if (wb.SheetNames.includes("Menu")) {
        const menuRows = XLSX.utils.sheet_to_json(wb.Sheets["Menu"]);
        if (menuRows.length > 0) {
          const parsedMenu = menuRows.map((row, idx) => ({
            id: Date.now() + idx,
            name: row.Producto || row.producto || row.Detalle || `Producto ${idx + 1}`,
            category: row.Categoria || row.categoria || "Minutas",
            price: parseFloat(row.Precio ?? row.precio ?? 0) || 0,
            stock: parseInt(row.Stock ?? row.stock ?? 30, 10) || 0,
            description: row.Descripción || row.Descripcion || "",
          }));
          setProducts(parsedMenu);
        }
      }

      // Si el Excel trae hoja "Inventario Of." o una hoja genérica con Detalle
      const invSheetName = wb.SheetNames.includes("Inventario Of.")
        ? "Inventario Of."
        : wb.SheetNames[0];
      const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[invSheetName], {
        defval: "",
      });

      const validSupplies = rawRows
        .filter(
          (r) =>
            r["Detalle"] &&
            String(r["Detalle"]).trim() !== "" &&
            String(r["Detalle"]).toLowerCase() !== "detalle"
        )
        .map((r, i) => ({
          id: `imp-${Date.now()}-${i}`,
          name: String(r["Detalle"]).trim(),
          type: r["Tipo"] ? String(r["Tipo"]).trim() : "Materia Prima",
          stock: parseFloat(r["Total disponible"]) || 0,
        }));

      if (validSupplies.length > 0) {
        setSupplies(validSupplies);
      }

      alert("Archivo Excel cargado correctamente.");
    };
    reader.readAsBinaryString(file);
  };

  const exportAllToExcel = () => {
    const wb = XLSX.utils.book_new();

    const menuData = products.map((p) => ({
      Producto: p.name,
      Categoria: p.category,
      Precio: p.price,
      Stock: p.stock,
      Descripcion: p.description,
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(menuData), "Menu");

    const invData = supplies.map((s) => ({
      Detalle: s.name,
      Tipo: s.type,
      "Total disponible": s.stock,
    }));
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(invData),
      "Inventario"
    );

    const salesData = sales.map((s) => ({
      Fecha: s.date,
      Productos: s.items.map((i) => `${i.qty}x ${i.name}`).join(", "),
      Total: Number(s.total.toFixed(2)),
      Metodo_Pago: s.paymentMethod,
      Efectivo_Recibido: Number(s.cashGiven.toFixed(2)),
      Cambio: Number(s.change.toFixed(2)),
    }));
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(salesData),
      "Ventas"
    );

    XLSX.writeFile(wb, "Minutas_el_Chele_Control.xlsx");
  };

  // --- CÁLCULOS FINANCIEROS ---
  const totalSold = sales.reduce((acc, s) => acc + s.total, 0);
  const totalDebts = financial.debts.reduce(
    (acc, d) => acc + (parseFloat(d.amount) || 0),
    0
  );
  const totalPurchases = financial.purchasesSummary.reduce(
    (acc, p) =>
      acc +
      (parseFloat(p.utensils) || 0) +
      (parseFloat(p.packaging) || 0) +
      (parseFloat(p.rawMaterial) || 0) +
      (parseFloat(p.transport) || 0),
    0
  );
  const totalPaidPurchases = financial.purchasesSummary.reduce(
    (acc, p) => acc + (parseFloat(p.paid) || 0),
    0
  );
  const pendingPurchases = totalPaidPurchases - totalPurchases;
  const netProfit = totalSold - totalDebts;
  const profitMinusReinvest = netProfit - (parseFloat(financial.reinvestment) || 0);

  const filteredProducts = products.filter((p) => {
    const matchesSearch = p.name.toLowerCase().includes(search.toLowerCase());
    const matchesCat =
      categoryFilter === "Todos" || p.category === categoryFilter;
    return matchesSearch && matchesCat;
  });

  return (
    <div className="min-h-screen flex flex-col bg-slate-100 text-slate-800">
      {/* MODAL DE CONTRASEÑA */}
      {pendingAction && (
        <div className="fixed inset-0 z-50 bg-black/60 flex items-center justify-center p-4">
          <form
            onSubmit={handlePasswordSubmit}
            className="bg-white rounded-2xl shadow-xl border border-slate-200 max-w-sm w-full p-6 space-y-4"
          >
            <div className="flex items-center gap-2 text-blue-700 font-bold text-lg">
              <Lock size={22} />
              <span>
                {pendingAction === "financial"
                  ? "Acceso a Financiero"
                  : "Modificar Precios"}
              </span>
            </div>
            <p className="text-sm text-slate-600">
              {pendingAction === "financial"
                ? "Esta sección está protegida. Ingresa la contraseña del dueño para ver y registrar datos financieros."
                : "Solo el dueño puede modificar precios. Ingresa la contraseña para desbloquear."}
            </p>
            <div>
              <input
                type="password"
                autoFocus
                placeholder="Contraseña"
                value={passwordInput}
                onChange={(e) => setPasswordInput(e.target.value)}
                className="w-full p-2.5 border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
              {passwordError && (
                <p className="text-xs text-red-600 font-medium mt-1.5">
                  {passwordError}
                </p>
              )}
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setPendingAction(null)}
                className="px-4 py-2 rounded-xl border border-slate-300 text-slate-700 text-sm font-medium hover:bg-slate-50"
              >
                Cancelar
              </button>
              <button
                type="submit"
                className="px-4 py-2 rounded-xl bg-blue-700 text-white text-sm font-semibold hover:bg-blue-800"
              >
                Desbloquear
              </button>
            </div>
          </form>
        </div>
      )}

      {/* ENCABEZADO */}
      <header className="bg-blue-800 text-white shadow-md">
        <div className="max-w-7xl mx-auto px-4 py-4 flex flex-wrap justify-between items-center gap-3">
          <div>
            <h1 className="text-xl sm:text-2xl font-bold tracking-tight">
              🍧 Minutas el Chele
            </h1>
            <p className="text-xs text-blue-200">
              Ventas, Inventario y Control de Negocio
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
              <ShoppingCart size={17} /> Ventas (POS)
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
              <History size={17} /> Historial ({sales.length})
            </button>

            <button
              onClick={handleSelectFinancialTab}
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

      {/* CONTENIDO PRINCIPAL */}
      <main className="max-w-7xl w-full mx-auto p-4 flex-1">
        {/* 1. PUNTO DE VENTA */}
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
                    placeholder="Buscar minuta o licuado..."
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
                    className={`p-4 rounded-2xl border text-left transition flex flex-col justify-between min-h-[140px] ${
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
                          product.stock > 10
                            ? "bg-green-100 text-green-800"
                            : product.stock > 0
                            ? "bg-amber-100 text-amber-800"
                            : "bg-red-100 text-red-800"
                        }`}
                      >
                        {product.stock > 0
                          ? `Disp: ${product.stock}`
                          : "Agotado"}
                      </span>
                    </div>
                  </button>
                ))}
              </div>
            </div>

            {/* CAJA DE COBRO */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-col justify-between h-fit sticky top-4">
              <div>
                <div className="flex items-center justify-between border-b pb-3 mb-3">
                  <h2 className="text-lg font-bold">Detalle de Venta</h2>
                  {cart.length > 0 && (
                    <button
                      onClick={() => setCart([])}
                      className="text-xs text-red-600 font-medium hover:underline"
                    >
                      Vaciar
                    </button>
                  )}
                </div>

                {cart.length === 0 ? (
                  <p className="text-slate-400 text-center py-8 text-sm">
                    Selecciona un producto para comenzar la venta.
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
                        Recibido ($):
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
                    actualizado!
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* 2. INVENTARIO Y PRECIOS PROTEGIDOS */}
        {activeTab === "inventory" && (
          <div className="space-y-6">
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
              <div>
                <h2 className="text-lg font-bold">
                  Inventario y Lista de Productos
                </h2>
                <p className="text-sm text-slate-500">
                  Sube tu Excel para actualizar productos o insumos. Los precios
                  están protegidos con contraseña.
                </p>
              </div>

              <div className="flex flex-wrap gap-2.5">
                <button
                  type="button"
                  onClick={() =>
                    pricesUnlocked
                      ? setPricesUnlocked(false)
                      : openPasswordModal("prices")
                  }
                  className={`flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-sm font-semibold border transition ${
                    pricesUnlocked
                      ? "bg-amber-100 border-amber-400 text-amber-900"
                      : "bg-slate-100 border-slate-300 text-slate-700 hover:bg-slate-200"
                  }`}
                >
                  {pricesUnlocked ? <Unlock size={16} /> : <Lock size={16} />}
                  {pricesUnlocked ? "Precios Desbloqueados" : "Desbloquear Precios"}
                </button>

                <label className="cursor-pointer flex items-center gap-1.5 bg-blue-700 hover:bg-blue-800 text-white px-4 py-2 rounded-xl text-sm font-semibold transition">
                  <Upload size={16} /> Cargar Excel
                  <input
                    type="file"
                    accept=".xlsx, .xls, .csv"
                    onChange={handleExcelUpload}
                    className="hidden"
                  />
                </label>

                <button
                  type="button"
                  onClick={exportAllToExcel}
                  className="flex items-center gap-1.5 bg-slate-800 hover:bg-slate-900 text-white px-4 py-2 rounded-xl text-sm font-semibold transition"
                >
                  <Download size={16} /> Descargar Excel
                </button>
              </div>
            </div>

            {/* Productos del Menú */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-4">
              <h3 className="font-bold text-slate-800">
                Productos a la Venta (Menú)
              </h3>

              <form
                onSubmit={handleAddProduct}
                className="flex flex-wrap gap-3 items-end bg-slate-50 p-3.5 rounded-xl border border-slate-200"
              >
                <div className="flex-1 min-w-[180px]">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">
                    Nombre del Producto
                  </label>
                  <input
                    type="text"
                    required
                    value={newProdName}
                    onChange={(e) => setNewProdName(e.target.value)}
                    placeholder="Ej. Minuta de tamarindo"
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
                    <option value="Bebidas">Bebidas</option>
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

                <div className="w-28">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">
                    Stock
                  </label>
                  <input
                    type="number"
                    required
                    value={newProdStock}
                    onChange={(e) => setNewProdStock(e.target.value)}
                    placeholder="25"
                    className="w-full p-2 border rounded-lg bg-white text-sm"
                  />
                </div>

                <button
                  type="submit"
                  className="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1"
                >
                  <Plus size={16} /> Agregar
                </button>
              </form>

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead className="bg-slate-50 border-b text-xs uppercase text-slate-500">
                    <tr>
                      <th className="p-3">Producto</th>
                      <th className="p-3">Categoría</th>
                      <th className="p-3">Precio ($)</th>
                      <th className="p-3">Stock</th>
                      <th className="p-3 text-right">Eliminar</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y text-sm">
                    {products.map((prod) => (
                      <tr key={prod.id} className="hover:bg-slate-50">
                        <td className="p-3 font-semibold">{prod.name}</td>
                        <td className="p-3 text-slate-600">{prod.category}</td>
                        <td className="p-3">
                          <div className="flex items-center gap-1.5">
                            <span className="text-slate-400">$</span>
                            <input
                              type="number"
                              step="0.01"
                              readOnly={!pricesUnlocked}
                              onClick={() => {
                                if (!pricesUnlocked) openPasswordModal("prices");
                              }}
                              value={prod.price}
                              onChange={(e) =>
                                handlePriceChange(prod.id, e.target.value)
                              }
                              className={`w-24 p-1.5 border rounded-lg font-bold ${
                                pricesUnlocked
                                  ? "bg-white border-amber-400 text-slate-900"
                                  : "bg-slate-100 border-slate-200 text-slate-700 cursor-pointer"
                              }`}
                            />
                            {!pricesUnlocked && (
                              <Lock size={14} className="text-slate-400" />
                            )}
                          </div>
                        </td>
                        <td className="p-3">
                          <input
                            type="number"
                            value={prod.stock}
                            onChange={(e) =>
                              handleProductStockChange(prod.id, e.target.value)
                            }
                            className="w-20 p-1.5 border rounded-lg font-semibold"
                          />
                        </td>
                        <td className="p-3 text-right">
                          <button
                            onClick={() =>
                              setProducts(
                                products.filter((item) => item.id !== prod.id)
                              )
                            }
                            className="text-red-600 hover:bg-red-50 p-1.5 rounded-lg"
                          >
                            <Trash2 size={16} />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Insumos (Empaques y Materia Prima de Inventario Of.) */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-4">
              <h3 className="font-bold text-slate-800">
                Insumos: Empaques y Materia Prima
              </h3>

              <form
                onSubmit={handleAddSupply}
                className="flex flex-wrap gap-3 items-end bg-slate-50 p-3.5 rounded-xl border border-slate-200"
              >
                <div className="flex-1 min-w-[200px]">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">
                    Detalle del Insumo
                  </label>
                  <input
                    type="text"
                    required
                    value={newSupplyName}
                    onChange={(e) => setNewSupplyName(e.target.value)}
                    placeholder="Ej. Vasos de 12oz"
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
                <div className="w-28">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">
                    Disponible
                  </label>
                  <input
                    type="number"
                    required
                    value={newSupplyStock}
                    onChange={(e) => setNewSupplyStock(e.target.value)}
                    placeholder="100"
                    className="w-full p-2 border rounded-lg bg-white text-sm"
                  />
                </div>
                <button
                  type="submit"
                  className="bg-blue-700 hover:bg-blue-800 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1"
                >
                  <Plus size={16} /> Añadir Insumo
                </button>
              </form>

              <div className="overflow-x-auto max-h-96">
                <table className="w-full text-left border-collapse">
                  <thead className="bg-slate-50 border-b text-xs uppercase text-slate-500 sticky top-0">
                    <tr>
                      <th className="p-3">Detalle</th>
                      <th className="p-3">Tipo</th>
                      <th className="p-3">Total disponible</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y text-sm">
                    {supplies.map((sup) => (
                      <tr key={sup.id} className="hover:bg-slate-50">
                        <td className="p-3 font-medium">{sup.name}</td>
                        <td className="p-3 text-slate-600">{sup.type}</td>
                        <td className="p-3">
                          <input
                            type="number"
                            value={sup.stock}
                            onChange={(e) => {
                              const val = parseFloat(e.target.value) || 0;
                              setSupplies(
                                supplies.map((item) =>
                                  item.id === sup.id
                                    ? { ...item, stock: val }
                                    : item
                                )
                              );
                            }}
                            className="w-24 p-1.5 border rounded-lg font-bold"
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* 3. HISTORIAL DE VENTAS */}
        {activeTab === "history" && (
          <div className="space-y-6">
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
              <div>
                <h2 className="text-lg font-bold">Registro Ordenado de Ventas</h2>
                <p className="text-2xl font-extrabold text-green-700 mt-1">
                  Total Vendido: ${totalSold.toFixed(2)}
                </p>
              </div>
              <button
                onClick={exportAllToExcel}
                className="flex items-center gap-2 bg-slate-800 hover:bg-slate-900 text-white px-4 py-2.5 rounded-xl text-sm font-semibold"
              >
                <Download size={17} /> Exportar a Excel
              </button>
            </div>

            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-x-auto">
              {sales.length === 0 ? (
                <p className="text-center text-slate-400 py-10">
                  Aún no se han registrado ventas.
                </p>
              ) : (
                <table className="w-full text-left border-collapse">
                  <thead className="bg-slate-50 border-b text-xs uppercase text-slate-500">
                    <tr>
                      <th className="p-4">Fecha y Hora</th>
                      <th className="p-4">Productos</th>
                      <th className="p-4">Pago</th>
                      <th className="p-4">Total</th>
                      <th className="p-4">Cambio</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y text-sm">
                    {sales.map((sale) => (
                      <tr key={sale.id} className="hover:bg-slate-50">
                        <td className="p-4 text-slate-500 whitespace-nowrap">
                          {sale.date}
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
                        </td>
                        <td className="p-4 font-extrabold text-slate-900">
                          ${sale.total.toFixed(2)}
                        </td>
                        <td className="p-4 text-slate-600">
                          {sale.paymentMethod === "Efectivo"
                            ? `$${sale.change.toFixed(2)}`
                            : "-"}
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
                  <h2 className="text-lg font-bold">Panel Financiero</h2>
                  <p className="text-xs text-slate-500">
                    Visible solo con contraseña del dueño
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
                <EyeOff size={16} /> Bloquear Financiero
              </button>
            </div>

            {/* Tarjetas Resumen */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  Meta
                </p>
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
                  className="mt-1 w-full text-xl font-extrabold text-blue-800 border-b border-dashed border-slate-300 focus:outline-none"
                />
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  Vendido
                </p>
                <p className="mt-1 text-xl font-extrabold text-green-700">
                  ${totalSold.toFixed(2)}
                </p>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  Deudas
                </p>
                <p className="mt-1 text-xl font-extrabold text-red-600">
                  ${totalDebts.toFixed(2)}
                </p>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  Ganancia
                </p>
                <p
                  className={`mt-1 text-xl font-extrabold ${
                    netProfit >= 0 ? "text-emerald-700" : "text-red-600"
                  }`}
                >
                  ${netProfit.toFixed(2)}
                </p>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  Reinversión
                </p>
                <input
                  type="number"
                  step="0.01"
                  value={financial.reinvestment}
                  onChange={(e) =>
                    setFinancial({
                      ...financial,
                      reinvestment: parseFloat(e.target.value) || 0,
                    })
                  }
                  className="mt-1 w-full text-xl font-extrabold text-slate-900 border-b border-dashed border-slate-300 focus:outline-none"
                />
                <p className="text-[11px] text-slate-500 mt-1">
                  Ganancia - reinver.: ${profitMinusReinvest.toFixed(2)}
                </p>
              </div>
            </div>

            {/* Compras y Pagos */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-3">
              <h3 className="font-bold text-slate-800">
                Compras por Tipo (Utensilios, Empaques, Materia Prima, Transporte)
              </h3>
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-sm">
                  <thead className="bg-slate-50 border-b text-xs uppercase text-slate-500">
                    <tr>
                      <th className="p-3">Tipo</th>
                      <th className="p-3">Utensilios</th>
                      <th className="p-3">Empaques</th>
                      <th className="p-3">Materia Prima</th>
                      <th className="p-3">Transporte</th>
                      <th className="p-3">Total</th>
                      <th className="p-3">Pagado/abono</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y">
                    {financial.purchasesSummary.map((row) => {
                      const rowTotal =
                        (parseFloat(row.utensils) || 0) +
                        (parseFloat(row.packaging) || 0) +
                        (parseFloat(row.rawMaterial) || 0) +
                        (parseFloat(row.transport) || 0);
                      return (
                        <tr key={row.id}>
                          <td className="p-3 font-semibold">{row.type}</td>
                          <td className="p-3">
                            ${(parseFloat(row.utensils) || 0).toFixed(2)}
                          </td>
                          <td className="p-3">
                            ${(parseFloat(row.packaging) || 0).toFixed(2)}
                          </td>
                          <td className="p-3">
                            ${(parseFloat(row.rawMaterial) || 0).toFixed(2)}
                          </td>
                          <td className="p-3">
                            ${(parseFloat(row.transport) || 0).toFixed(2)}
                          </td>
                          <td className="p-3 font-bold">${rowTotal.toFixed(2)}</td>
                          <td className="p-3 text-green-700 font-semibold">
                            ${(parseFloat(row.paid) || 0).toFixed(2)}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <p className="text-xs text-slate-600 pt-1">
                Total compras: <strong>${totalPurchases.toFixed(2)}</strong> |
                Total abonado: <strong>${totalPaidPurchases.toFixed(2)}</strong> |
                Diferencia: <strong>${pendingPurchases.toFixed(2)}</strong>
              </p>
            </div>

            {/* Deudas */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-3">
              <h3 className="font-bold text-slate-800">
                Deudas y Compromisos (Total: ${totalDebts.toFixed(2)})
              </h3>
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3">
                {financial.debts.map((debt) => (
                  <div
                    key={debt.id}
                    className="p-3.5 rounded-xl border border-slate-200 bg-slate-50"
                  >
                    <p className="text-sm font-bold text-slate-800">
                      {debt.concept}
                    </p>
                    {debt.note && (
                      <p className="text-xs text-slate-500">{debt.note}</p>
                    )}
                    <input
                      type="number"
                      step="0.01"
                      value={debt.amount}
                      onChange={(e) => {
                        const val = parseFloat(e.target.value) || 0;
                        setFinancial({
                          ...financial,
                          debts: financial.debts.map((d) =>
                            d.id === debt.id ? { ...d, amount: val } : d
                          ),
                        });
                      }}
                      className="mt-2 w-full p-1.5 bg-white border rounded-lg font-bold text-red-700 text-sm"
                    />
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}