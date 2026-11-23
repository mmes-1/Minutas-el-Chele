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
  EyeOff,
  ClipboardCheck,
  Calendar,
  AlertCircle,
  X,
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

// Inventario inicial tomado de "Inventario Of." del Excel
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
    name: "agua alpina en botella 699ml",
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
  const [saleSuccess, setSaleSuccess] = useState(false);

  // Contraseña para Financiero y Precios
  const [pricesUnlocked, setPricesUnlocked] = useState(false);
  const [financialUnlocked, setFinancialUnlocked] = useState(false);
  const [passwordInput, setPasswordInput] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [pendingAction, setPendingAction] = useState(null); // 'prices' | 'financial'

  // Modal "Actualizar inventario"
  const [updateInvModalOpen, setUpdateInvModalOpen] = useState(false);
  const [invUpdateDate, setInvUpdateDate] = useState(
    new Date().toISOString().slice(0, 10)
  );
  const [invUpdateMode, setInvUpdateMode] = useState("entrada"); // 'entrada' | 'salida' | 'conteo'
  const [invUpdateNote, setInvUpdateNote] = useState("");
  const [invDraftInputs, setInvDraftInputs] = useState({});

  // Nuevo producto de menú
  const [newProdName, setNewProdName] = useState("");
  const [newProdCategory, setNewProdCategory] = useState("Minutas");
  const [newProdPrice, setNewProdPrice] = useState("");
  const [newProdStock, setNewProdStock] = useState("");

  // Nuevos registros financieros (deuda / reinversión / fecha de pago)
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
    const savedProducts = localStorage.getItem("chele_v3_products");
    const savedSupplies = localStorage.getItem("chele_v3_supplies");
    const savedInvLogs = localStorage.getItem("chele_v3_inv_logs");
    const savedSales = localStorage.getItem("chele_v3_sales");
    const savedFinancial = localStorage.getItem("chele_v3_financial");

    setProducts(savedProducts ? JSON.parse(savedProducts) : INITIAL_MENU);
    setSupplies(savedSupplies ? JSON.parse(savedSupplies) : INITIAL_SUPPLIES);
    setInventoryLogs(savedInvLogs ? JSON.parse(savedInvLogs) : []);
    setSales(savedSales ? JSON.parse(savedSales) : []);
    setFinancial(
      savedFinancial ? JSON.parse(savedFinancial) : INITIAL_FINANCIAL
    );
  }, []);

  useEffect(() => {
    if (products.length > 0) {
      localStorage.setItem("chele_v3_products", JSON.stringify(products));
    }
  }, [products]);

  useEffect(() => {
    if (supplies.length > 0) {
      localStorage.setItem("chele_v3_supplies", JSON.stringify(supplies));
    }
  }, [supplies]);

  useEffect(() => {
    localStorage.setItem("chele_v3_inv_logs", JSON.stringify(inventoryLogs));
  }, [inventoryLogs]);

  useEffect(() => {
    localStorage.setItem("chele_v3_sales", JSON.stringify(sales));
  }, [sales]);

  useEffect(() => {
    localStorage.setItem("chele_v3_financial", JSON.stringify(financial));
  }, [financial]);

  // --- CONTRASEÑA ---
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
      setPasswordError("Contraseña incorrecta. Ingresa Negocio2026.");
    }
  };

  // --- ACTUALIZAR INVENTARIO CON HISTORIAL ---
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
      alert("El monto en efectivo debe ser igual o mayor al total.");
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

  // --- PRECIOS PROTEGIDOS ---
  const handlePriceChange = (productId, rawValue) => {
    if (!pricesUnlocked) {
      openPasswordModal("prices");
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

  const handleAddProduct = (e) => {
    e.preventDefault();
    if (!newProdName.trim()) return;
    const priceVal = parseFloat(newProdPrice);
    if (!pricesUnlocked && priceVal > 0) {
      openPasswordModal("prices");
      return;
    }
    setProducts([
      ...products,
      {
        id: Date.now(),
        name: newProdName.trim(),
        category: newProdCategory,
        price: Number.isNaN(priceVal) ? 0 : priceVal,
        stock: parseInt(newProdStock, 10) || 0,
        description: "",
      },
    ]);
    setNewProdName("");
    setNewProdPrice("");
    setNewProdStock("");
  };

  // --- FINANCIERO: EMPLEADOS, DEUDAS Y REINVERSIONES ---
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
          Cambio: Number(s.change.toFixed(2)),
        }))
      ),
      "Ventas"
    );

    XLSX.writeFile(wb, "Minutas_el_Chele_Completo.xlsx");
  };

  // --- CÁLCULOS FINANCIEROS ---
  const totalSold = sales.reduce((acc, s) => acc + s.total, 0);

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

  const totalReinvestments = financial.reinvestments.reduce(
    (sum, r) => sum + (parseFloat(r.amount) || 0),
    0
  );

  // Dinero disponible real en caja (ventas menos pagos ya realizados y reinversiones)
  const availableCash = totalSold - paidEmployeePay - totalReinvestments;
  const totalPendingObligations = pendingEmployeePay + pendingOtherDebts;

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
            <div className="flex items-center gap-2 text-blue-800 font-bold text-lg">
              <Lock size={20} />
              <span>
                {pendingAction === "financial"
                  ? "Acceso a Financiero"
                  : "Modificar Precios"}
              </span>
            </div>
            <p className="text-sm text-slate-600">
              Ingresa la contraseña del dueño para continuar.
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
                onClick={() => setPendingAction(null)}
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
                          product.stock > 10
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

                <button
                  onClick={handleCompleteSale}
                  disabled={cart.length === 0}
                  className="w-full py-3 bg-green-600 hover:bg-green-700 disabled:bg-slate-300 text-white font-bold rounded-xl shadow transition"
                >
                  Guardar Venta
                </button>

                {saleSuccess && (
                  <div className="flex items-center justify-center gap-2 text-green-800 bg-green-100 p-2.5 rounded-xl text-sm font-semibold">
                    <CheckCircle2 size={18} /> ¡Venta guardada!
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* 2. INVENTARIO DIARIO Y PRECIOS PROTEGIDOS */}
        {activeTab === "inventory" && (
          <div className="space-y-6">
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
              <div>
                <h2 className="text-lg font-bold">
                  Inventario Diario y Control de Stock
                </h2>
                <p className="text-sm text-slate-500">
                  Inicia con tu inventario del Excel. Para evitar errores
                  manuales, el stock de insumos solo cambia usando{" "}
                  <strong>Actualizar inventario</strong>.
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
                  {pricesUnlocked
                    ? "Precios Desbloqueados"
                    : "Desbloquear Precios"}
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
                Productos del Menú (Precios protegidos con Negocio2026)
              </h3>

              <form
                onSubmit={handleAddProduct}
                className="flex flex-wrap gap-3 items-end bg-slate-50 p-3.5 rounded-xl border border-slate-200"
              >
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
                <div className="w-28">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">
                    Stock
                  </label>
                  <input
                    type="number"
                    required
                    value={newProdStock}
                    onChange={(e) => setNewProdStock(e.target.value)}
                    placeholder="20"
                    className="w-full p-2 border rounded-lg bg-white text-sm"
                  />
                </div>
                <button
                  type="submit"
                  className="bg-blue-700 hover:bg-blue-800 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1"
                >
                  <Plus size={16} /> Agregar
                </button>
              </form>

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-sm">
                  <thead className="bg-slate-50 border-b text-xs uppercase text-slate-500">
                    <tr>
                      <th className="p-3">Producto</th>
                      <th className="p-3">Categoría</th>
                      <th className="p-3">Precio ($)</th>
                      <th className="p-3">Stock</th>
                      <th className="p-3 text-right">Eliminar</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y">
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
                                  ? "bg-white border-amber-400"
                                  : "bg-slate-100 border-slate-200 cursor-pointer"
                              }`}
                            />
                            {!pricesUnlocked && (
                              <Lock size={14} className="text-slate-400" />
                            )}
                          </div>
                        </td>
                        <td className="p-3 font-bold">{prod.stock}</td>
                        <td className="p-3 text-right">
                          <button
                            onClick={() =>
                              setProducts(
                                products.filter((p) => p.id !== prod.id)
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

            {/* Materiales e Insumos (Stock bloqueado para edición directa) */}
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-4">
              <div className="flex flex-wrap justify-between items-center gap-2">
                <div>
                  <h3 className="font-bold text-slate-800">
                    Materiales e Insumos (Empaques y Materia Prima)
                  </h3>
                  <p className="text-xs text-slate-500">
                    Stock protegido contra edición directa. Usa{" "}
                    <strong>Actualizar inventario</strong> arriba para modificarlo.
                  </p>
                </div>
                <span className="inline-flex items-center gap-1 text-xs font-semibold bg-slate-100 text-slate-700 px-3 py-1 rounded-full">
                  <Lock size={13} /> Stock protegido
                </span>
              </div>

              <div className="overflow-x-auto max-h-96">
                <table className="w-full text-left border-collapse text-sm">
                  <thead className="bg-slate-50 border-b text-xs uppercase text-slate-500 sticky top-0">
                    <tr>
                      <th className="p-3">Detalle</th>
                      <th className="p-3">Tipo</th>
                      <th className="p-3">Total disponible</th>
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

        {/* 3. HISTORIAL DE VENTAS */}
        {activeTab === "history" && (
          <div className="space-y-6">
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
              <div>
                <h2 className="text-lg font-bold">Historial de Ventas</h2>
                <p className="text-2xl font-extrabold text-green-700 mt-1">
                  Total Vendido: ${totalSold.toFixed(2)}
                </p>
              </div>
              <button
                onClick={exportToExcel}
                className="flex items-center gap-2 bg-slate-800 hover:bg-slate-900 text-white px-4 py-2.5 rounded-xl text-sm font-semibold"
              >
                <Download size={17} /> Exportar Reporte Excel
              </button>
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
                    </tr>
                  </thead>
                  <tbody className="divide-y">
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
                  <h2 className="text-lg font-bold">Control Financiero</h2>
                  <p className="text-xs text-slate-500">
                    Disponibilidad de pagos, fechas de Kat y Gab, deudas y
                    reinversiones
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

            {/* Indicadores clave */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  Total Vendido
                </p>
                <p className="mt-1 text-2xl font-extrabold text-green-700">
                  ${totalSold.toFixed(2)}
                </p>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  Efectivo Disponible para Pagar
                </p>
                <p
                  className={`mt-1 text-2xl font-extrabold ${
                    availableCash > 0 ? "text-emerald-700" : "text-slate-600"
                  }`}
                >
                  ${availableCash.toFixed(2)}
                </p>
                <p className="text-[11px] text-slate-500 mt-1">
                  {availableCash > 0
                    ? "Hay dinero disponible para cubrir pagos"
                    : "Aún no hay saldo disponible de ventas"}
                </p>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  Pendiente por Pagar
                </p>
                <p className="mt-1 text-2xl font-extrabold text-red-600">
                  ${totalPendingObligations.toFixed(2)}
                </p>
                <p className="text-[11px] text-slate-500 mt-1">
                  Salarios: ${pendingEmployeePay.toFixed(2)} | Deudas: $
                  {pendingOtherDebts.toFixed(2)}
                </p>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200">
                <p className="text-xs font-bold uppercase text-slate-500">
                  Reinversión Registrada
                </p>
                <p className="mt-1 text-2xl font-extrabold text-blue-800">
                  ${totalReinvestments.toFixed(2)}
                </p>
                <p className="text-[11px] text-slate-500 mt-1">
                  Meta: ${Number(financial.goal || 0).toFixed(2)}
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

              {/* Agregar nueva fecha de pago para Kat o Gab */}
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
              {/* Deudas */}
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
                      !debt.paid && availableCash >= debt.amount && debt.amount > 0;
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

              {/* Reinversiones */}
              <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-4">
                <h3 className="font-bold text-slate-900">
                  Registro de Reinversiones
                </h3>

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
                        <span className="font-extrabold text-blue-800">
                          ${Number(reinv.amount).toFixed(2)}
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