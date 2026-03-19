import axios from "axios";
import { Clock, Download, Edit, Factory, LogOut, RefreshCw, Trash2, Users, Wrench } from "lucide-react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { BrowserRouter, Navigate, Route, Routes } from "react-router-dom";
import { toast } from "sonner";
import * as XLSX from 'xlsx';
import "./App.css";
import { Badge } from "./components/ui/badge";
import { Button } from "./components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "./components/ui/card";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "./components/ui/dialog";
import { Input } from "./components/ui/input";
import { Label } from "./components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "./components/ui/select";
import { Toaster } from "./components/ui/sonner";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "./components/ui/tabs";
import { Textarea } from "./components/ui/textarea";

const BACKEND_URL = process.env.REACT_APP_BACKEND_URL;
const API = `${BACKEND_URL}/api`;

const App = () => {
  const [user, setUser] = useState(null);
  const [token, setToken] = useState(localStorage.getItem("token"));
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (token) {
      validateToken();
    } else {
      setLoading(false);
    }
  }, [token]);

  const validateToken = async () => {
    try {
      const response = await axios.get(`${API}/auth/me`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      setUser(response.data);
    } catch (error) {
      localStorage.removeItem("token");
      setToken(null);
    }
    setLoading(false);
  };

  const login = async (username, password) => {
    try {
      const response = await axios.post(`${API}/auth/login`, { username, password });
      const { token: newToken, user: userData } = response.data;
      localStorage.setItem("token", newToken);
      setToken(newToken);
      setUser(userData);
      toast.success("Login realizado com sucesso!");
    } catch (error) {
      toast.error("Credenciais inválidas");
    }
  };

  const logout = () => {
    localStorage.removeItem("token");
    setToken(null);
    setUser(null);
    toast.success("Logout realizado com sucesso!");
  };

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center login-container">
        <div className="loading-spinner"></div>
      </div>
    );
  }

  if (!user) {
    return <LoginPage onLogin={login} />;
  }

  return (
    <div className="min-h-screen">
      <Toaster />
      <BrowserRouter>
        <Routes>
          <Route path="/" element={<Dashboard user={user} onLogout={logout} />} />
          <Route path="*" element={<Navigate to="/" replace />} />
        </Routes>
      </BrowserRouter>
    </div>
  );
};

const LoginPage = ({ onLogin }) => {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [isLoading, setIsLoading] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setIsLoading(true);
    await onLogin(username, password);
    setIsLoading(false);
  };

  return (
    <div className="login-container">
      <Card className="w-full max-w-md login-card">
        <CardHeader className="text-center pb-8">
          <h1 className="text-5xl font-bold text-white mb-4">MercoTêxtil</h1>
          <p className="login-subtitle">Sistema de Controle de Máquinas</p>
        </CardHeader>
        <CardContent className="form-merco">
          <form onSubmit={handleSubmit} className="space-y-4">
            <div className="space-y-2">
              <Label htmlFor="username">Usuário</Label>
              <Input
                id="username"
                type="text"
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                placeholder="Digite seu usuário"
                required
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="password">Senha</Label>
              <Input
                id="password"
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Digite sua senha"
                required
              />
            </div>
            <Button type="submit" className="w-full btn-merco" disabled={isLoading}>
              {isLoading ? "Entrando..." : "Entrar"}
            </Button>
          </form>
        </CardContent>
      </Card>
    </div>
  );
};

const Dashboard = ({ user, onLogout }) => {
  const [activeLayout, setActiveLayout] = useState("16_fusos");
  const [machines, setMachines] = useState([]);
  const [orders, setOrders] = useState([]);
  const [users, setUsers] = useState([]);
  const [maintenances, setMaintenances] = useState([]);
  const [espulas, setEspulas] = useState([]);

  useEffect(() => {
    loadData();
    const interval = setInterval(loadData, 5000); // Auto-refresh every 5 seconds
    return () => clearInterval(interval);
  }, [activeLayout, user.role]);

  const loadData = async () => {
    await Promise.all([
      loadMachines(),
      loadOrders(),
      loadMaintenances(),
      loadEspulas(),
      user.role === "admin" ? loadUsers() : Promise.resolve()
    ]);
  };

  const loadMachines = async () => {
    try {
      const response = await axios.get(`${API}/machines/${activeLayout}`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      setMachines(response.data);
    } catch (error) {
      toast.error("Erro ao carregar máquinas");
    }
  };

  const loadOrders = async () => {
    try {
      const response = await axios.get(`${API}/orders`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      setOrders(response.data);
    } catch (error) {
      toast.error("Erro ao carregar pedidos");
    }
  };

  const loadUsers = async () => {
    try {
      const response = await axios.get(`${API}/users`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      setUsers(response.data);
    } catch (error) {
      toast.error("Erro ao carregar usuários");
    }
  };

  const loadMaintenances = async () => {
    try {
      const response = await axios.get(`${API}/maintenance`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      setMaintenances(response.data);
    } catch (error) {
      toast.error("Erro ao carregar manutenções");
    }
  };

  const loadEspulas = async () => {
    try {
      const response = await axios.get(`${API}/espulas`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      setEspulas(response.data);
    } catch (error) {
      toast.error("Erro ao carregar espulagem");
    }
  };

  const resetDatabase = async () => {
    if (window.confirm("Tem certeza que deseja resetar o banco de dados? Isso apagará todos os pedidos, espulagem e manutenções.")) {
      try {
        await axios.post(`${API}/reset-database`, {}, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        });
        
        toast.success("Banco de dados resetado com sucesso!");
        loadData();
      } catch (error) {
        toast.error("Erro ao resetar banco de dados");
      }
    }
  };

  const getRoleBadge = (role) => {
    switch (role) {
      case "admin": return { text: "Administrador", class: "badge-admin" };
      case "operador_interno": return { text: "Op. Interno", class: "badge-interno" };
      case "operador_externo": return { text: "Op. Externo", class: "badge-externo" };
      default: return { text: role, class: "badge-merco" };
    }
  };

  const badge = getRoleBadge(user.role);

  return (
    <div className="min-h-screen">
      {/* Header */}
      <header className="merco-header">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-20">
            <div className="flex items-center space-x-6">
              <span className="text-3xl font-bold text-white">MercoTêxtil</span>
              <Badge className={`${badge.class} badge-merco`}>
                {badge.text}
              </Badge>
            </div>
            <div className="flex items-center space-x-4">
              <span className="text-sm text-gray-400">Olá, <span className="text-white font-medium">{user.username}</span></span>
              {user.role === "admin" && (
                <Button variant="outline" size="sm" onClick={resetDatabase} className="border-yellow-600 text-yellow-400 hover:bg-yellow-600 hover:text-white">
                  <RefreshCw className="h-4 w-4 mr-2" />
                  Reset DB
                </Button>
              )}
              <Button variant="outline" size="sm" onClick={onLogout} className="border-red-600 text-red-400 hover:bg-red-600 hover:text-white">
                <LogOut className="h-4 w-4 mr-2" />
                Sair
              </Button>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <Tabs defaultValue="dashboard" className="space-y-6">
          <TabsList className="tabs-merco">
            {user.permissions?.dashboard !== false && <TabsTrigger value="dashboard" className="tab-merco">Dashboard</TabsTrigger>}
            {user.permissions?.producao !== false && <TabsTrigger value="orders" className="tab-merco">Produção</TabsTrigger>}
            {user.permissions?.ordem_producao !== false && <TabsTrigger value="ordemproducao" className="tab-merco">Ordem de Produção</TabsTrigger>}
            {user.permissions?.relatorios !== false && <TabsTrigger value="relatorios" className="tab-merco">Relatórios</TabsTrigger>}
            {user.permissions?.espulagem !== false && <TabsTrigger value="espulas" className="tab-merco">Espulagem</TabsTrigger>}
            {user.permissions?.manutencao !== false && <TabsTrigger value="maintenance" className="tab-merco">Manutenção</TabsTrigger>}
            {user.permissions?.banco_dados !== false && <TabsTrigger value="bancodados" className="tab-merco">Banco de Dados</TabsTrigger>}
            {(user.role === "admin" || user.permissions?.administracao === true) && <TabsTrigger value="admin" className="tab-merco">Administração</TabsTrigger>}
          </TabsList>

          <TabsContent value="dashboard" className="space-y-6">
            <div className="flex justify-between items-center">
              <h2 className="text-3xl font-bold text-white">Painel de Controle</h2>
              <div className="layout-switcher">
                <button
                  className={`layout-btn ${activeLayout === "16_fusos" ? "active" : ""}`}
                  onClick={() => setActiveLayout("16_fusos")}
                >
                  16 Fusos
                </button>
                <button
                  className={`layout-btn ${activeLayout === "32_fusos" ? "active" : ""}`}
                  onClick={() => setActiveLayout("32_fusos")}
                >
                  32 Fusos
                </button>
              </div>
            </div>

            <FusosPanel 
              layout={activeLayout} 
              machines={machines} 
              user={user}
              onMachineUpdate={loadMachines}
              onOrderUpdate={loadOrders}
              onMaintenanceUpdate={loadMaintenances}
            />
          </TabsContent>

          <TabsContent value="orders">
            <OrdersPanel orders={orders} user={user} onOrderUpdate={loadOrders} onMachineUpdate={loadMachines} />
          </TabsContent>

          <TabsContent value="ordemproducao">
            <OrdemProducaoPanel user={user} />
          </TabsContent>

          <TabsContent value="relatorios">
            <RelatoriosPanel user={user} />
          </TabsContent>

          <TabsContent value="espulas">
            <EspulasPanel espulas={espulas} user={user} onEspulaUpdate={loadEspulas} />
          </TabsContent>

          <TabsContent value="maintenance">
            <MaintenancePanel maintenances={maintenances} user={user} onMaintenanceUpdate={loadMaintenances} onMachineUpdate={loadMachines} />
          </TabsContent>

          {user.permissions?.banco_dados !== false && (
            <TabsContent value="bancodados">
              <BancoDadosPanel user={user} />
            </TabsContent>
          )}

          {(user.role === "admin" || user.permissions?.administracao === true) && (
            <TabsContent value="admin">
              <AdminPanel users={users} onUserUpdate={loadUsers} />
            </TabsContent>
          )}
        </Tabs>

        {/* Footer */}
        <footer className="mt-16 pt-8 border-t border-gray-700 text-center">
          <p className="text-gray-400 text-sm">
            © 2026 MercoTêxtil - Todos os direitos reservados | Desenvolvido por CodeliumCompany
          </p>
        </footer>
      </main>
    </div>
  );
};

const FusosPanel = ({ layout, machines, user, onMachineUpdate, onOrderUpdate, onMaintenanceUpdate }) => {
  const [selectedMachine, setSelectedMachine] = useState(null);
  const [maintenanceMachine, setMaintenanceMachine] = useState(null);
  const [showQueue, setShowQueue] = useState(false);
  const [queueMachine, setQueueMachine] = useState(null);
  const [machineOrders, setMachineOrders] = useState([]);
  const [showManualOrder, setShowManualOrder] = useState(false);
  const [manualOrderMachine, setManualOrderMachine] = useState(null);
  const [orderData, setOrderData] = useState({
    cliente: "",
    artigo: "",
    cor: "",
    quantidade: "",
    observacao: ""
  });
  const [maintenanceData, setMaintenanceData] = useState({
    motivo: ""
  });

  // Função para obter cor do status - incluindo desativada
  const getStatusColor = (status) => {
    switch (status) {
      case "verde": return "status-verde";
      case "amarelo": return "status-amarelo";
      case "vermelho": return "status-vermelho";
      case "azul": return "status-azul";
      case "desativada": return "status-desativada";
      default: return "status-verde";
    }
  };

  // Função para admin ativar/desativar máquina
  const toggleMachineActive = async (machineId) => {
    if (user?.role !== "admin") {
      toast.error("Apenas administradores podem ativar/desativar máquinas");
      return;
    }
    
    try {
      const response = await axios.put(`${API}/machines/${machineId}/toggle-active`, {}, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success(response.data.message);
      onMachineUpdate(); // Refresh machines
    } catch (error) {
      toast.error("Erro ao alterar status da máquina");
    }
  };

  // CORRIGIR FORMATAÇÃO DE HORÁRIO - CONVERSÃO CORRETA UTC PARA BRASÍLIA
  const formatDateTimeBrazil = (utcString) => {
    if (!utcString) return "-";
    try {
      // Parse UTC string e converte para horário de Brasília
      const utcDate = new Date(utcString + (utcString.includes('Z') ? '' : 'Z'));
      
      return new Intl.DateTimeFormat('pt-BR', {
        timeZone: 'America/Sao_Paulo',
        day: '2-digit',
        month: '2-digit', 
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false
      }).format(utcDate);
    } catch (error) {
      console.error('Error formatting date:', error);
      return "-";
    }
  };

  const loadMachineQueue = async (machineCode) => {
    try {
      const response = await axios.get(`${API}/machines/${machineCode}/orders`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      setMachineOrders(response.data);
    } catch (error) {
      console.error("Erro ao carregar fila:", error);
      toast.error("Erro ao carregar fila de pedidos");
    }
  };

  const handleMachineClick = (machine) => {
    if (user.role === "admin" || user.role === "operador_interno") {
      if (machine.status === "verde") {
        setSelectedMachine(machine);
      } else if (machine.status === "amarelo" || machine.status === "vermelho") {
        // Máquina com pedidos (pendentes ou em produção) - mostrar fila
        setQueueMachine(machine);
        loadMachineQueue(machine.code);
        setShowQueue(true);
      }
    } else if (user.role === "operador_externo") {
      // Operador externo pode ver fila de pedidos pendentes ou em produção
      if (machine.status === "amarelo" || machine.status === "vermelho") {
        setQueueMachine(machine);
        loadMachineQueue(machine.code);
        setShowQueue(true);
      }
    }
  };

  const startOrderProduction = async (orderId, machineCode) => {
    try {
      await axios.put(`${API}/machines/${machineCode}/orders/${orderId}/start`, {}, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Produção iniciada com sucesso!");
      loadMachineQueue(machineCode);
      onMachineUpdate();
      onOrderUpdate();
    } catch (error) {
      toast.error(error.response?.data?.detail || "Erro ao iniciar produção");
    }
  };

  const finishOrderProduction = async (orderId, machineCode) => {
    try {
      await axios.put(`${API}/machines/${machineCode}/orders/${orderId}/finish`, {}, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Pedido finalizado com sucesso!");
      loadMachineQueue(machineCode);
      onMachineUpdate();
      onOrderUpdate();
    } catch (error) {
      toast.error(error.response?.data?.detail || "Erro ao finalizar pedido");
    }
  };

  const openManualOrderDialog = (machine) => {
    setManualOrderMachine(machine);
    setOrderData({
      cliente: "",
      artigo: "",
      cor: "",
      quantidade: "",
      observacao: ""
    });
    setShowManualOrder(true);
  };

  const createManualOrder = async () => {
    try {
      await axios.post(`${API}/machines/${manualOrderMachine.code}/orders`, orderData, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Pedido criado na máquina com sucesso!");
      setShowManualOrder(false);
      onMachineUpdate();
      onOrderUpdate();
    } catch (error) {
      toast.error(error.response?.data?.detail || "Erro ao criar pedido");
    }
  };

  const handleMaintenanceClick = (machine) => {
    // Permite manutenção em qualquer status exceto se já estiver em manutenção ou desativada
    if (machine.status !== "azul" && machine.status !== "desativada") {
      setMaintenanceMachine(machine);
    } else if (machine.status === "azul") {
      toast.info("Esta máquina já está em manutenção");
    }
  };

  const handleOrderSubmit = async () => {
    try {
      await axios.post(`${API}/orders`, {
        machine_id: selectedMachine.id,
        ...orderData
      }, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Pedido criado com sucesso!");
      setSelectedMachine(null);
      setOrderData({ cliente: "", artigo: "", cor: "", quantidade: "", observacao: "" });
      onMachineUpdate();
      onOrderUpdate();
    } catch (error) {
      toast.error("Erro ao criar pedido");
    }
  };

  const handleMaintenanceSubmit = async () => {
    try {
      await axios.post(`${API}/maintenance`, {
        machine_id: maintenanceMachine.id,
        motivo: maintenanceData.motivo
      }, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Máquina colocada em manutenção!");
      setMaintenanceMachine(null);
      setMaintenanceData({ motivo: "" });
      onMachineUpdate();
      onMaintenanceUpdate();
    } catch (error) {
      toast.error("Erro ao colocar máquina em manutenção");
    }
  };

  const exportMapaTrancadeiras = async (layoutType) => {
    try {
      const XLSXStyle = require('xlsx-js-style');
      const response = await axios.get(
        `${API}/reports/mapa-trancadeiras?layout_type=${layoutType}`,
        { headers: { Authorization: `Bearer ${localStorage.getItem("token")}` } }
      );

      const { machines: machineData } = response.data;
      const machinesMap = {};
      machineData.forEach(m => { machinesMap[m.code] = m; });

      const getMachineData = (code) => {
        const m = machinesMap[code];
        if (!m) return [code, '', '', '', ''];
        if (m.status === 'azul') return [code, '', '', '', 'MANUTENÇÃO'];
        if (m.status === 'desativada') return [code, '', '', '', 'DESATIVADA'];
        if (m.status === 'vermelho' && m.order) {
          return [code, m.order.artigo || '', m.order.cor || '', m.order.quantidade_metros || '', 'EM PRODUÇÃO'];
        }
        if (m.status === 'amarelo' && m.order) {
          return [code, m.order.artigo || '', m.order.cor || '', m.order.quantidade_metros || '', 'PENDENTE'];
        }
        return [code, '', '', '', ''];
      };

      const thin = { style: "thin", color: { rgb: "000000" } };
      const bdrAll = { top: thin, bottom: thin, left: thin, right: thin };
      const centerWrap = { horizontal: "center", vertical: "center", wrapText: true };

      const colLetter = (idx) => {
        if (idx < 26) return String.fromCharCode(65 + idx);
        return String.fromCharCode(65 + Math.floor(idx / 26) - 1) + String.fromCharCode(65 + (idx % 26));
      };

      const wb = XLSXStyle.utils.book_new();

      if (layoutType === '32_fusos') {
        // ── 32 FUSOS: direct cell approach with merge boxes ──
        // Row layout (0-indexed):
        //  r  0- 2:  CT1[0]…CT12[11]
        //  r  3:     GAP
        //  r  4- 6:  CT13[0]…CT24[11]
        //  r  7:     GAP
        //  r  8-10:  U1[1] U2[2]  |  U11[5] U12[6]  |  U21[9] U22[10]
        //  r 11-13:  U3[1] U4[2]  |  U13[5] U14[6]  |  U23[9] U24[10]
        //  r 14-16:  U5[1] U6[2]  |  U15[5] U16[6]  |  U25[9] U26[10]
        //  r 17-19:  U7[1] U8[2]  |  U17[5] U18[6]  |  U27[9] U28[10]
        //  r 20-22:  U9[1] U10[2] |  U19[5] U20[6]  |  U29[9] U30[10]
        //  r 23:     GAP
        //  r 24-26:  N1[1]…N10[10]
        //  r 27:     GAP
        //  r 28-30:  U32[4]  U31[6]
        //  r 31-33:  U33[4]
        //  r 34:     Date [col 9]

        const ws32 = {};
        const TOTAL_ROWS_32 = 55;
        const numCols32 = 12; // A-L

        const sc32 = (r, c, val, style) => {
          const addr = `${colLetter(c)}${r + 1}`;
          ws32[addr] = { v: String(val), t: 's', s: style };
        };

        const box5_32 = (code, r, c) => {
          const [codeVal, artigo, cor, qtd, statusLabel] = getMachineData(code);
          const lines = [codeVal];
          if (artigo) lines.push(artigo);
          if (cor) lines.push(cor);
          if (qtd) lines.push(qtd);
          if (statusLabel) lines.push(statusLabel);
          sc32(r,   c, lines.join('\n'), { border: bdrAll, alignment: centerWrap, font: { sz: 9, bold: true } });
          sc32(r+1, c, '', { border: bdrAll, alignment: centerWrap });
          sc32(r+2, c, '', { border: bdrAll, alignment: centerWrap });
          sc32(r+3, c, '', { border: bdrAll, alignment: centerWrap });
          sc32(r+4, c, '', { border: bdrAll, alignment: centerWrap });
        };

        // ── CT rows ──
        ['CT1','CT2','CT3','CT4','CT5','CT6','CT7','CT8','CT9','CT10','CT11','CT12'].forEach((code, i) => {
          box5_32(code, 0, i);
        });
        ['CT13','CT14','CT15','CT16','CT17','CT18','CT19','CT20','CT21','CT22','CT23','CT24'].forEach((code, i) => {
          box5_32(code, 6, i);
        });

        // ── U groups (3 columns of pairs) ──
        const uPairs32 = [
          ['U1','U2',   'U11','U12',  'U21','U22',  12],
          ['U3','U4',   'U13','U14',  'U23','U24',  17],
          ['U5','U6',   'U15','U16',  'U25','U26',  22],
          ['U7','U8',   'U17','U18',  'U27','U28',  27],
          ['U9','U10',  'U19','U20',  'U29','U30',  32],
        ];
        uPairs32.forEach(([a, b, c, d, e, f, startRow]) => {
          box5_32(a, startRow, 1); box5_32(b, startRow, 2);
          box5_32(c, startRow, 5); box5_32(d, startRow, 6);
          box5_32(e, startRow, 9); box5_32(f, startRow, 10);
        });

        // ── N row ──
        ['N1','N2','N3','N4','N5','N6','N7','N8','N9','N10'].forEach((code, i) => {
          box5_32(code, 38, i + 1);
        });

        // ── U31-U33 ──
        box5_32('U32', 44, 4);
        box5_32('U31', 44, 6);
        box5_32('U33', 49, 4);

        // ── Date ──
        sc32(54, 9, new Date().toLocaleDateString('pt-BR'), {
          font: { sz: 10 }, alignment: { horizontal: "center" }
        });

        ws32['!ref'] = `A1:${colLetter(numCols32 - 1)}${TOTAL_ROWS_32}`;

        // ── Merges (5-row single-column boxes) ──
        ws32['!merges'] = [];
        const m32 = [];
        // CT rows
        for (let i = 0; i < 12; i++) { m32.push([0, i]); m32.push([6, i]); }
        // U pairs
        uPairs32.forEach(([,,,,, , startRow]) => {
          [1,2,5,6,9,10].forEach(c => m32.push([startRow, c]));
        });
        // N row
        for (let i = 1; i <= 10; i++) m32.push([38, i]);
        // U31-U33
        m32.push([44, 4]); m32.push([44, 6]);
        m32.push([49, 4]);
        m32.forEach(([r, c]) => ws32['!merges'].push({ s: { r, c }, e: { r: r+4, c } }));

        // ── Column widths ──
        ws32['!cols'] = Array(numCols32).fill({ wch: 14 });

        // ── Row heights ──
        ws32['!rows'] = Array.from({ length: TOTAL_ROWS_32 }, () => ({ hpt: 15 }));
        ws32['!rows'][5]  = { hpt: 6 };
        ws32['!rows'][11] = { hpt: 6 };
        ws32['!rows'][37] = { hpt: 6 };
        ws32['!rows'][43] = { hpt: 6 };
        ws32['!rows'][54] = { hpt: 16 };

        XLSXStyle.utils.book_append_sheet(wb, ws32, 'Mapa 32 Fusos');

      } else {
        // ── 16 FUSOS: direct cell approach with merge boxes ──
        // Row layout (0-indexed):
        //  r 0-2:   CD1[1] CD2[2] | CD5[4] CD6[5] | CD17[7] CD18[8] CD19[9] CD20[10]
        //  r 3-5:   CD3[1] CD4[2] | CD7[4] CD8[5] | CD21[7] CD22[8] CD23[9] CD24[10]
        //  r 6:     GAP
        //  r 7-9:   CD9[1] CD10[2] | CD13[4] CD14[5]
        //  r 10-12: CD11[1] CD12[2] | CD15[4] CD16[5]
        //  r 13:    GAP
        //  r 14:    "17 FUSOS" label (cols 4-7 merged)
        //  r 15-17: CI1[4] CI2[5] CI3[6] CI4[7]
        //  r 18:    GAP
        //  r 19-21: F23[1] F21[2] … F1[12]
        //  r 22-24: F24[1] F22[2] … F2[12]
        //  r 25:    Date [col 10]

        const ws = {};
        const TOTAL_ROWS = 40;
        const numCols16 = 13; // A-M

        const sc = (r, c, val, style) => {
          const addr = `${colLetter(c)}${r + 1}`;
          ws[addr] = { v: String(val), t: 's', s: style };
        };

        // Standard 5-row merge box: content in anchor, blanks for rows 1-4
        const box5 = (code, r, c) => {
          const [codeVal, artigo, cor, qtd, statusLabel] = getMachineData(code);
          const lines = [codeVal];
          if (artigo) lines.push(artigo);
          if (cor) lines.push(cor);
          if (qtd) lines.push(qtd);
          if (statusLabel) lines.push(statusLabel);
          sc(r,   c, lines.join('\n'), { border: bdrAll, alignment: centerWrap, font: { sz: 9, bold: true } });
          sc(r+1, c, '', { border: bdrAll, alignment: centerWrap });
          sc(r+2, c, '', { border: bdrAll, alignment: centerWrap });
          sc(r+3, c, '', { border: bdrAll, alignment: centerWrap });
          sc(r+4, c, '', { border: bdrAll, alignment: centerWrap });
        };

        // ── Top section ──
        box5('CD1', 0, 1); box5('CD2', 0, 2);
        box5('CD5', 0, 4); box5('CD6', 0, 5);
        box5('CD17', 0, 7); box5('CD18', 0, 8); box5('CD19', 0, 9); box5('CD20', 0, 10);
        box5('CD3', 5, 1); box5('CD4', 5, 2);
        box5('CD7', 5, 4); box5('CD8', 5, 5);
        box5('CD21', 5, 7); box5('CD22', 5, 8); box5('CD23', 5, 9); box5('CD24', 5, 10);

        // ── Middle section ──
        box5('CD9',  11, 1); box5('CD10', 11, 2);
        box5('CD13', 11, 4); box5('CD14', 11, 5);
        box5('CD11', 16, 1); box5('CD12', 16, 2);
        box5('CD15', 16, 4); box5('CD16', 16, 5);

        // ── "17 FUSOS" label row (r=22, cols 4-7 merged) ──
        sc(22, 4, '17 FUSOS', { border: bdrAll, alignment: { horizontal: "center", vertical: "center" }, font: { sz: 12, bold: true } });
        [5, 6, 7].forEach(c => sc(22, c, '', { border: bdrAll, alignment: centerWrap }));

        // ── CI block (r=23-27, cols 4-7) ──
        box5('CI1', 23, 4); box5('CI2', 23, 5); box5('CI3', 23, 6); box5('CI4', 23, 7);

        // ── F machines row 1: F23,F21,...,F1 (r=29-33, cols 1-12) ──
        ['F23','F21','F19','F17','F15','F13','F11','F9','F7','F5','F3','F1'].forEach((code, i) => {
          box5(code, 29, i + 1);
        });

        // ── F machines row 2: F24,F22,...,F2 (r=34-38, cols 1-12) ──
        ['F24','F22','F20','F18','F16','F14','F12','F10','F8','F6','F4','F2'].forEach((code, i) => {
          box5(code, 34, i + 1);
        });

        // ── Date ──
        sc(39, 10, new Date().toLocaleDateString('pt-BR'), {
          font: { sz: 10 }, alignment: { horizontal: "center" }
        });

        ws['!ref'] = `A1:${colLetter(numCols16 - 1)}${TOTAL_ROWS}`;

        // ── Merges ──
        ws['!merges'] = [];

        // Standard 5-row single-column merges
        const m5 = [
          [0,1],[0,2],[0,4],[0,5],[0,7],[0,8],[0,9],[0,10],          // CD1,CD2,CD5,CD6,CD17-CD20
          [5,1],[5,2],[5,4],[5,5],[5,7],[5,8],[5,9],[5,10],           // CD3,CD4,CD7,CD8,CD21-CD24
          [11,1],[11,2],[11,4],[11,5],                                 // CD9,CD10,CD13,CD14
          [16,1],[16,2],[16,4],[16,5],                                 // CD11,CD12,CD15,CD16
          [23,4],[23,5],[23,6],[23,7],                                 // CI1,CI2,CI3,CI4
        ];
        for (let i = 1; i <= 12; i++) { m5.push([29, i]); m5.push([34, i]); } // F rows
        m5.forEach(([r, c]) => ws['!merges'].push({ s: { r, c }, e: { r: r+4, c } }));

        // "17 FUSOS" horizontal label merge (r=22, cols 4-7)
        ws['!merges'].push({ s: { r: 22, c: 4 }, e: { r: 22, c: 7 } });

        // ── Column widths ──
        ws['!cols'] = [{ wch: 2 }, ...Array(12).fill({ wch: 14 })]; // A=margin, B-M=14

        // ── Row heights ──
        ws['!rows'] = Array.from({ length: TOTAL_ROWS }, () => ({ hpt: 15 }));
        ws['!rows'][10] = { hpt: 6 };  // gap between top and middle sections
        ws['!rows'][21] = { hpt: 6 };  // gap before 17 FUSOS label
        ws['!rows'][22] = { hpt: 18 }; // "17 FUSOS" label row
        ws['!rows'][28] = { hpt: 6 };  // gap before F machines
        ws['!rows'][39] = { hpt: 16 }; // date row

        XLSXStyle.utils.book_append_sheet(wb, ws, 'Mapa 16 Fusos');
      }

      const layoutLabel = layoutType === '32_fusos' ? '32_fusos' : '16_fusos';
      const dateStr = new Date().toISOString().split('T')[0];
      XLSXStyle.writeFile(wb, `mapa_trancadeiras_${layoutLabel}_${dateStr}.xlsx`);
      toast.success("Relatório exportado com sucesso!");
    } catch (error) {
      console.error("Erro ao exportar mapa:", error);
      toast.error("Erro ao exportar mapa das trançadeiras");
    }
  };

  // 16 Fusos Layout - EXACT replication with stable keys
  const renderLayout16 = () => {
    const machinesByCode = useMemo(() => {
      const result = {};
      machines.forEach(machine => {
        if (machine?.code) {
          result[machine.code] = machine;
        }
      });
      return result;
    }, [machines]);

    const renderMachineBox = useCallback((machine, code) => {
      const uniqueKey = machine?.id ? `machine-16-${machine.id}` : `empty-16-${code}`;
      
      return (
        <div key={uniqueKey} className={`machine-box-16 ${getStatusColor(machine?.status || 'verde')}`}>
          <span onClick={() => machine && machine.status !== 'desativada' && handleMachineClick(machine)} className="machine-code">
            {machine?.code || code}
          </span>
          {machine && (
            <>
              <button className="maintenance-btn" onClick={(e) => {
                e.stopPropagation();
                handleMaintenanceClick(machine);
              }}>
                <Wrench className="h-4 w-4" />
              </button>
              {user?.role === "admin" && (
                <button 
                  className="admin-toggle-btn" 
                  onClick={(e) => {
                    e.stopPropagation();
                    toggleMachineActive(machine.id);
                  }}
                  title={machine.status === 'desativada' ? 'Reativar máquina' : 'Desativar máquina'}
                >
                  {machine.status === 'desativada' ? '✓' : '✕'}
                </button>
              )}
            </>
          )}
        </div>
      );
    }, [machines, getStatusColor, handleMachineClick, handleMaintenanceClick, toggleMachineActive, user]);

    return (
      <div className="layout-16-exact-centered">
        {/* Top section - CD blocks horizontally arranged */}
        <div className="layout-16-top-section">
          {/* CD1-CD4 block (2x2) */}
          <div className="cd-block-2x2">
            <div className="cd-row">
              {renderMachineBox(machinesByCode["CD1"], "CD1")}
              {renderMachineBox(machinesByCode["CD2"], "CD2")}
            </div>
            <div className="cd-row">
              {renderMachineBox(machinesByCode["CD3"], "CD3")}
              {renderMachineBox(machinesByCode["CD4"], "CD4")}
            </div>
          </div>
          
          {/* CD5-CD8 block (2x2) */}
          <div className="cd-block-2x2">
            <div className="cd-row">
              {renderMachineBox(machinesByCode["CD5"], "CD5")}
              {renderMachineBox(machinesByCode["CD6"], "CD6")}
            </div>
            <div className="cd-row">
              {renderMachineBox(machinesByCode["CD7"], "CD7")}
              {renderMachineBox(machinesByCode["CD8"], "CD8")}
            </div>
          </div>
          
          {/* CD17-CD20 block (horizontal line) */}
          <div className="cd-block-horizontal">
            {renderMachineBox(machinesByCode["CD17"], "CD17")}
            {renderMachineBox(machinesByCode["CD18"], "CD18")}
            {renderMachineBox(machinesByCode["CD19"], "CD19")}
            {renderMachineBox(machinesByCode["CD20"], "CD20")}
          </div>
        </div>

        {/* Middle section - CD blocks horizontally arranged */}
        <div className="layout-16-middle-section">
          {/* CD9-CD12 block (2x2) */}
          <div className="cd-block-2x2">
            <div className="cd-row">
              {renderMachineBox(machinesByCode["CD9"], "CD9")}
              {renderMachineBox(machinesByCode["CD10"], "CD10")}
            </div>
            <div className="cd-row">
              {renderMachineBox(machinesByCode["CD11"], "CD11")}
              {renderMachineBox(machinesByCode["CD12"], "CD12")}
            </div>
          </div>
          
          {/* CD13-CD16 block (2x2) */}
          <div className="cd-block-2x2">
            <div className="cd-row">
              {renderMachineBox(machinesByCode["CD13"], "CD13")}
              {renderMachineBox(machinesByCode["CD14"], "CD14")}
            </div>
            <div className="cd-row">
              {renderMachineBox(machinesByCode["CD15"], "CD15")}
              {renderMachineBox(machinesByCode["CD16"], "CD16")}
            </div>
          </div>
          
          {/* CD21-CD24 block (horizontal line) */}
          <div className="cd-block-horizontal">
            {renderMachineBox(machinesByCode["CD21"], "CD21")}
            {renderMachineBox(machinesByCode["CD22"], "CD22")}
            {renderMachineBox(machinesByCode["CD23"], "CD23")}
            {renderMachineBox(machinesByCode["CD24"], "CD24")}
          </div>
        </div>

        {/* CI section - "17 FUSOS" */}
        <div className="layout-16-ci-section">
          <div className="ci-label">17 FUSOS</div>
          <div className="ci-block-horizontal">
            {renderMachineBox(machinesByCode["CI1"], "CI1")}
            {renderMachineBox(machinesByCode["CI2"], "CI2")}
            {renderMachineBox(machinesByCode["CI3"], "CI3")}
            {renderMachineBox(machinesByCode["CI4"], "CI4")}
          </div>
        </div>

        {/* F section - Bottom rows */}
        <div className="layout-16-f-section">
          <div className="f-row-container">
            <div className="f-row">
              {["F23", "F21", "F19", "F17", "F15", "F13", "F11", "F9", "F7", "F5", "F3", "F1"].map(code => 
                renderMachineBox(machinesByCode[code], code)
              )}
            </div>
            <div className="f-row">
              {["F24", "F22", "F20", "F18", "F16", "F14", "F12", "F10", "F8", "F6", "F4", "F2"].map(code => 
                renderMachineBox(machinesByCode[code], code)
              )}
            </div>
          </div>
        </div>
      </div>
    );
  };

  // 32 Fusos Layout - EXACT replication with stable keys  
  const renderLayout32 = () => {
    const machinesByCode = useMemo(() => {
      const result = {};
      machines.forEach(machine => {
        if (machine?.code) {
          result[machine.code] = machine;
        }
      });
      return result;
    }, [machines]);

    const renderMachineBox = useCallback((machine, code) => {
      const uniqueKey = machine?.id ? `machine-32-${machine.id}` : `empty-32-${code}`;
      
      return (
        <div key={uniqueKey} className={`machine-box-32 ${getStatusColor(machine?.status || 'verde')}`}>
          <span onClick={() => machine && machine.status !== 'desativada' && handleMachineClick(machine)} className="machine-code">
            {machine?.code || code}
          </span>
          {machine && (
            <>
              <button className="maintenance-btn" onClick={(e) => {
                e.stopPropagation();
                handleMaintenanceClick(machine);
              }}>
                <Wrench className="h-4 w-4" />
              </button>
              {user?.role === "admin" && (
                <button 
                  className="admin-toggle-btn" 
                  onClick={(e) => {
                    e.stopPropagation();
                    toggleMachineActive(machine.id);
                  }}
                  title={machine.status === 'desativada' ? 'Reativar máquina' : 'Desativar máquina'}
                >
                  {machine.status === 'desativada' ? '✓' : '✕'}
                </button>
              )}
            </>
          )}
        </div>
      );
    }, [machines, getStatusColor, handleMachineClick, handleMaintenanceClick, toggleMachineActive, user]);

    return (
      <div className="layout-32-exact-new">
        {/* CT Machines - Row 1: CT1-CT12 */}
        <div className="ct-row-1">
          {["CT1", "CT2", "CT3", "CT4", "CT5", "CT6", "CT7", "CT8", "CT9", "CT10", "CT11", "CT12"].map(code => 
            renderMachineBox(machinesByCode[code], code)
          )}
        </div>

        {/* CT Machines - Row 2: CT13-CT24 */}
        <div className="ct-row-2">
          {["CT13", "CT14", "CT15", "CT16", "CT17", "CT18", "CT19", "CT20", "CT21", "CT22", "CT23", "CT24"].map(code => 
            renderMachineBox(machinesByCode[code], code)
          )}
        </div>

        {/* U Machines - Three blocks of 10 machines each (2x5 grid) */}
        <div className="u-machines-container">
          {/* Left Block: U1-U10 */}
          <div className="u-block">
            <div className="u-column">
              {["U1", "U3", "U5", "U7", "U9"].map(code => 
                renderMachineBox(machinesByCode[code], code)
              )}
            </div>
            <div className="u-column">
              {["U2", "U4", "U6", "U8", "U10"].map(code => 
                renderMachineBox(machinesByCode[code], code)
              )}
            </div>
          </div>

          {/* Middle Block: U11-U20 */}
          <div className="u-block">
            <div className="u-column">
              {["U11", "U13", "U15", "U17", "U19"].map(code => 
                renderMachineBox(machinesByCode[code], code)
              )}
            </div>
            <div className="u-column">
              {["U12", "U14", "U16", "U18", "U20"].map(code => 
                renderMachineBox(machinesByCode[code], code)
              )}
            </div>
          </div>

          {/* Right Block: U21-U30 */}
          <div className="u-block">
            <div className="u-column">
              {["U21", "U23", "U25", "U27", "U29"].map(code => 
                renderMachineBox(machinesByCode[code], code)
              )}
            </div>
            <div className="u-column">
              {["U22", "U24", "U26", "U28", "U30"].map(code => 
                renderMachineBox(machinesByCode[code], code)
              )}
            </div>
          </div>
        </div>

        {/* N Machines - Single row: N1-N10 */}
        <div className="n-machines-row">
          <div className="n-machines-confirmed">
            {["N1", "N2", "N3", "N4", "N5", "N6"].map(code => 
              renderMachineBox(machinesByCode[code], code)
            )}
          </div>
          <div className="n-machines-dashed">
            {["N7", "N8", "N9", "N10"].map(code => 
              renderMachineBox(machinesByCode[code], code)
            )}
          </div>
        </div>

        {/* Bottom U Machines: U31-U33 */}
        <div className="u-bottom-group">
          <div className="u-stack">
            {renderMachineBox(machinesByCode["U32"], "U32")}
            {renderMachineBox(machinesByCode["U33"], "U33")}
          </div>
          <div className="u-single">
            {renderMachineBox(machinesByCode["U31"], "U31")}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="space-y-6">
      <div className="fusos-container card-merco">
        <div className="mb-6">
          <div className="flex justify-between items-center mb-6">
            <h3 className="text-2xl font-bold text-white">
              Layout {layout === "16_fusos" ? "16" : "32"} Fusos
            </h3>
            {(user.role === "admin" || user.role === "operador_interno") && (
              <div className="flex gap-2">
                <Button
                  onClick={() => exportMapaTrancadeiras(layout)}
                  className="bg-green-600 hover:bg-green-700 text-white"
                >
                  <Download className="h-4 w-4 mr-2" />
                  Exportar Relatório
                </Button>
                <Button 
                  onClick={() => {
                    const availableMachines = machines.filter(m => m.status === "verde" || m.status === "amarelo");
                    if (availableMachines.length > 0) {
                      openManualOrderDialog(availableMachines[0]);
                    } else {
                      toast.error("Todas as máquinas estão em uso ou manutenção");
                    }
                  }} 
                  className="bg-blue-600 hover:bg-blue-700 text-white"
                >
                  + Lançar Pedido Manualmente
                </Button>
              </div>
            )}
          </div>
          <div className="status-legend">
            <div className="status-item">
              <div className="status-dot verde"></div>
              <span>Livre</span>
            </div>
            <div className="status-item">
              <div className="status-dot amarelo"></div>
              <span>Pendente</span>
            </div>
            <div className="status-item">
              <div className="status-dot vermelho"></div>
              <span>Em Uso</span>
            </div>
            <div className="status-item">
              <div className="status-dot azul"></div>
              <span>Manutenção</span>
            </div>
          </div>
        </div>
        
        {layout === "16_fusos" ? renderLayout16() : renderLayout32()}
      </div>

      {/* Order Dialog */}
      <Dialog open={!!selectedMachine} onOpenChange={() => setSelectedMachine(null)}>
        <DialogContent className="dialog-merco max-w-md">
          <DialogHeader className="dialog-header">
            <DialogTitle className="dialog-title">
              Novo Pedido - {selectedMachine?.code}
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-4 p-6 form-merco">
            <div>
              <Label htmlFor="cliente">Cliente *</Label>
              <Input
                id="cliente"
                value={orderData.cliente}
                onChange={(e) => setOrderData({...orderData, cliente: e.target.value})}
                placeholder="Nome do cliente"
                required
              />
            </div>
            <div>
              <Label htmlFor="artigo">Artigo *</Label>
              <Input
                id="artigo"
                value={orderData.artigo}
                onChange={(e) => setOrderData({...orderData, artigo: e.target.value})}
                placeholder="Artigo"
                required
              />
            </div>
            <div>
              <Label htmlFor="cor">Cor *</Label>
              <Input
                id="cor"
                value={orderData.cor}
                onChange={(e) => setOrderData({...orderData, cor: e.target.value})}
                placeholder="Cor"
                required
              />
            </div>
            <div>
              <Label htmlFor="quantidade">Quantidade *</Label>
              <Input
                id="quantidade"
                type="text"
                value={orderData.quantidade}
                onChange={(e) => setOrderData({...orderData, quantidade: e.target.value})}
                placeholder="Ex: 100, 50kg, 30m"
                required
              />
            </div>
            <div>
              <Label htmlFor="observacao">Observação</Label>
              <Textarea
                id="observacao"
                value={orderData.observacao}
                onChange={(e) => setOrderData({...orderData, observacao: e.target.value})}
                placeholder="Observações (opcional)"
              />
            </div>
            <Button 
              onClick={handleOrderSubmit} 
              className="w-full btn-merco"
              disabled={!orderData.cliente || !orderData.artigo || !orderData.cor || !orderData.quantidade}
            >
              Criar Pedido
            </Button>
          </div>
        </DialogContent>
      </Dialog>

      {/* Maintenance Dialog */}
      <Dialog open={!!maintenanceMachine} onOpenChange={() => setMaintenanceMachine(null)}>
        <DialogContent className="dialog-merco max-w-md">
          <DialogHeader className="dialog-header">
            <DialogTitle className="dialog-title">
              Manutenção - {maintenanceMachine?.code}
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-4 p-6 form-merco">
            <div>
              <Label htmlFor="motivo">Motivo da Manutenção *</Label>
              <Textarea
                id="motivo"
                value={maintenanceData.motivo}
                onChange={(e) => setMaintenanceData({...maintenanceData, motivo: e.target.value})}
                placeholder="Descreva o motivo da manutenção"
                required
              />
            </div>
            <Button 
              onClick={handleMaintenanceSubmit} 
              className="w-full btn-merco"
              disabled={!maintenanceData.motivo}
            >
              Colocar em Manutenção
            </Button>
          </div>
        </DialogContent>
      </Dialog>

      {/* Dialog de Fila de Pedidos */}
      <Dialog open={showQueue} onOpenChange={setShowQueue}>
        <DialogContent className="dialog-merco max-w-4xl">
          <DialogHeader className="dialog-header">
            <DialogTitle className="dialog-title">
              Fila de Pedidos - Máquina {queueMachine?.code}
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-4 p-6 max-h-[70vh] overflow-y-auto">
            {machineOrders.length === 0 ? (
              <p className="text-center text-gray-400 py-8">Nenhum pedido na fila</p>
            ) : (
              <div className="space-y-3">
                {machineOrders.map((order, index) => (
                  <Card key={order.id} className="card-merco">
                    <CardContent className="p-4">
                      <div className="flex items-center justify-between">
                        <div className="flex-1">
                          <div className="flex items-center gap-3 mb-2">
                            <span className="text-white font-bold text-lg">#{index + 1}</span>
                            {order.numero_os && (
                              <Badge className="bg-blue-600">OS {order.numero_os}</Badge>
                            )}
                            <Badge className={order.origem === "manual" ? "bg-purple-600" : "bg-green-600"}>
                              {order.origem === "manual" ? "Manual" : "Espulagem"}
                            </Badge>
                            <Badge className={
                              order.status === "pendente" ? "bg-yellow-600" :
                              order.status === "em_producao" ? "bg-red-600" : "bg-green-600"
                            }>
                              {order.status === "pendente" ? "Pendente" : 
                               order.status === "em_producao" ? "Em Produção" : "Finalizado"}
                            </Badge>
                          </div>
                          <div className="grid grid-cols-4 gap-3 text-sm">
                            <div>
                              <span className="text-gray-400">Cliente:</span>
                              <p className="text-white font-medium">{order.cliente}</p>
                            </div>
                            <div>
                              <span className="text-gray-400">Artigo:</span>
                              <p className="text-white font-medium">{order.artigo}</p>
                            </div>
                            <div>
                              <span className="text-gray-400">Cor:</span>
                              <p className="text-white font-medium">{order.cor}</p>
                            </div>
                            <div>
                              <span className="text-gray-400">Quantidade:</span>
                              <p className="text-white font-medium">{order.quantidade}</p>
                            </div>
                          </div>
                        </div>
                        <div className="flex gap-2 ml-4">
                          {order.status === "pendente" && (user.role === "operador_externo" || user.role === "admin" || user.role === "operador_interno") && (
                            <Button
                              onClick={() => startOrderProduction(order.id, queueMachine.code)}
                              className="bg-green-600 hover:bg-green-700"
                            >
                              Iniciar
                            </Button>
                          )}
                          {order.status === "em_producao" && (
                            <Button
                              onClick={() => finishOrderProduction(order.id, queueMachine.code)}
                              className="bg-blue-600 hover:bg-blue-700"
                            >
                              Finalizar
                            </Button>
                          )}
                        </div>
                      </div>
                    </CardContent>
                  </Card>
                ))}
              </div>
            )}
          </div>
        </DialogContent>
      </Dialog>

      {/* Dialog de Lançamento Manual */}
      <Dialog open={showManualOrder} onOpenChange={setShowManualOrder}>
        <DialogContent className="dialog-merco max-w-2xl">
          <DialogHeader className="dialog-header">
            <DialogTitle className="dialog-title">
              Lançar Pedido Manualmente - {manualOrderMachine?.code}
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-4 p-6 form-merco">
            <div>
              <Label>Selecione a Máquina *</Label>
              <Select 
                value={manualOrderMachine?.code || ""} 
                onValueChange={(value) => {
                  const machine = machines.find(m => m.code === value);
                  setManualOrderMachine(machine);
                }}
              >
                <SelectTrigger>
                  <SelectValue placeholder="Selecione a máquina" />
                </SelectTrigger>
                <SelectContent>
                  {machines.filter(m => m.status === "verde" || m.status === "amarelo").map((machine) => (
                    <SelectItem key={machine.id} value={machine.code}>
                      {machine.code} - {machine.layout_type === '16_fusos' ? '16 Fusos' : '32 Fusos'}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <Label htmlFor="cliente">Cliente *</Label>
                <Input
                  id="cliente"
                  value={orderData.cliente}
                  onChange={(e) => setOrderData({...orderData, cliente: e.target.value})}
                  placeholder="Nome do cliente"
                />
              </div>
              <div>
                <Label htmlFor="artigo">Artigo *</Label>
                <Input
                  id="artigo"
                  value={orderData.artigo}
                  onChange={(e) => setOrderData({...orderData, artigo: e.target.value})}
                  placeholder="Artigo"
                />
              </div>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <Label htmlFor="cor">Cor *</Label>
                <Input
                  id="cor"
                  value={orderData.cor}
                  onChange={(e) => setOrderData({...orderData, cor: e.target.value})}
                  placeholder="Cor"
                />
              </div>
              <div>
                <Label htmlFor="quantidade">Quantidade *</Label>
                <Input
                  id="quantidade"
                  value={orderData.quantidade}
                  onChange={(e) => setOrderData({...orderData, quantidade: e.target.value})}
                  placeholder="Quantidade"
                />
              </div>
            </div>
            <div>
              <Label htmlFor="observacao">Observação</Label>
              <Textarea
                id="observacao"
                value={orderData.observacao}
                onChange={(e) => setOrderData({...orderData, observacao: e.target.value})}
                placeholder="Observações do pedido"
              />
            </div>
            <Button 
              onClick={createManualOrder} 
              className="w-full btn-merco"
              disabled={!orderData.cliente || !orderData.artigo || !orderData.cor || !orderData.quantidade || !manualOrderMachine}
            >
              Criar Pedido na Máquina
            </Button>
          </div>
        </DialogContent>
      </Dialog>
    </div>
  );
};

const OrdersPanel = ({ orders, user, onOrderUpdate, onMachineUpdate }) => {
  const [selectedOrder, setSelectedOrder] = useState(null);
  const [finishData, setFinishData] = useState({
    observacao_liberacao: "",
    laudo_final: ""
  });

  const startOrderFromList = async (order) => {
    try {
      await axios.put(`${API}/machines/${order.machine_code}/orders/${order.id}/start`, {}, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Produção iniciada com sucesso!");
      onOrderUpdate();
      onMachineUpdate();
    } catch (error) {
      toast.error(error.response?.data?.detail || "Erro ao iniciar produção");
    }
  };

  const finishOrderFromList = async (order) => {
    try {
      await axios.put(`${API}/machines/${order.machine_code}/orders/${order.id}/finish`, {}, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Pedido finalizado com sucesso!");
      onOrderUpdate();
      onMachineUpdate();
    } catch (error) {
      toast.error(error.response?.data?.detail || "Erro ao finalizar pedido");
    }
  };

  const deleteOrder = async (order) => {
    if (window.confirm(`Tem certeza que deseja excluir o pedido ${order.client || 'sem cliente'}?`)) {
      try {
        await axios.delete(`${API}/orders/${order.id}`, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        });
        
        toast.success("Pedido excluído com sucesso!");
        onOrderUpdate();
        onMachineUpdate();
      } catch (error) {
        toast.error(error.response?.data?.detail || "Erro ao excluir pedido");
      }
    }
  };


  const updateOrder = async (orderId, status, observacao = "", laudo = "") => {
    try {
      await axios.put(`${API}/orders/${orderId}`, {
        status,
        observacao_liberacao: observacao,
        laudo_final: laudo
      }, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Pedido atualizado com sucesso!");
      onOrderUpdate();
      onMachineUpdate();
    } catch (error) {
      toast.error("Erro ao atualizar pedido");
    }
  };

  const handleFinish = async () => {
    await updateOrder(selectedOrder.id, "finalizado", finishData.observacao_liberacao, finishData.laudo_final);
    setSelectedOrder(null);
    setFinishData({ observacao_liberacao: "", laudo_final: "" });
  };

  const getStatusBadge = (status) => {
    switch (status) {
      case "pendente": return "bg-yellow-600 text-yellow-100";
      case "em_producao": return "bg-red-600 text-red-100";
      case "finalizado": return "bg-green-600 text-green-100";
      default: return "bg-gray-600 text-gray-100";
    }
  };

  // CORRIGIR FORMATAÇÃO DE HORÁRIO - CONVERSÃO CORRETA UTC PARA BRASÍLIA  
  const formatDateTimeBrazil = (utcString) => {
    if (!utcString) return "-";
    try {
      // Parse UTC string e converte para horário de Brasília
      const utcDate = new Date(utcString + (utcString.includes('Z') ? '' : 'Z'));
      
      return new Intl.DateTimeFormat('pt-BR', {
        timeZone: 'America/Sao_Paulo',
        day: '2-digit',
        month: '2-digit', 
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false
      }).format(utcDate);
    } catch (error) {
      console.error('Error formatting date:', error);
      return "-";
    }
  };

  const formatDateBrazil = (utcString) => {
    if (!utcString) return "-";
    try {
      // Parse UTC string e converte para horário de Brasília
      const utcDate = new Date(utcString + (utcString.includes('Z') ? '' : 'Z'));
      
      return new Intl.DateTimeFormat('pt-BR', {
        timeZone: 'America/Sao_Paulo',
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      }).format(utcDate);
    } catch (error) {
      console.error('Error formatting date:', error);
      return "-";
    }
  };

  const formatTime = (dateString) => {
    if (!dateString) return "-";
    try {
      // Assume que o backend sempre envia UTC
      const date = new Date(dateString);
      return new Intl.DateTimeFormat('pt-BR', {
        timeZone: 'America/Sao_Paulo',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false
      }).format(date);
    } catch (error) {
      return "-";
    }
  };

  const getCurrentBrazilTime = () => {
    const now = new Date();
    return new Intl.DateTimeFormat('pt-BR', {
      timeZone: 'America/Sao_Paulo',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      hour12: false
    }).format(now).replace(',', '');
  };

  return (
    <div className="space-y-6">
      <h2 className="text-3xl font-bold text-white">Gerenciamento de Pedidos</h2>
      <div className="grid gap-6">
        {orders.map((order) => (
          <Card key={order.id} className="card-merco-large">
            <CardContent className="p-6">
              <div className="flex justify-between items-start mb-6">
                <div>
                  <h3 className="font-bold text-white text-xl">
                    Máquina {order.machine_code} - {order.layout_type.replace('_', ' ')}
                  </h3>
                  <p className="text-gray-400 text-lg">Cliente: <span className="text-white">{order.cliente}</span></p>
                </div>
                <Badge className={`${getStatusBadge(order.status)} font-semibold text-sm`}>
                  {order.status.replace("_", " ").toUpperCase()}
                </Badge>
              </div>
              
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-base mb-6">
                <div>
                  <span className="font-medium text-gray-400">Artigo:</span>
                  <p className="text-white">{order.artigo}</p>
                </div>
                <div>
                  <span className="font-medium text-gray-400">Cor:</span>
                  <p className="text-white">{order.cor}</p>
                </div>
                <div>
                  <span className="font-medium text-gray-400">Quantidade:</span>
                  <p className="text-white">{order.quantidade}</p>
                </div>
                <div>
                  <span className="font-medium text-gray-400">Criado por:</span>
                  <p className="text-white">{order.created_by}</p>
                </div>
              </div>

              {/* Time information */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-base mb-6 p-4 bg-gray-900/50 rounded-lg">
                <div className="flex items-center space-x-3">
                  <Clock className="h-5 w-5 text-blue-400" />
                  <div>
                    <span className="font-medium text-gray-400">Criado:</span>
                    <p className="text-white">{formatDateTimeBrazil(order.created_at)}</p>
                  </div>
                </div>
                <div className="flex items-center space-x-3">
                  <Clock className="h-5 w-5 text-orange-400" />
                  <div>
                    <span className="font-medium text-gray-400">Iniciado:</span>
                    <p className="text-white">{formatDateTimeBrazil(order.started_at)}</p>
                  </div>
                </div>
                <div className="flex items-center space-x-3">
                  <Clock className="h-5 w-5 text-green-400" />
                  <div>
                    <span className="font-medium text-gray-400">Finalizado:</span>
                    <p className="text-white">{formatDateTimeBrazil(order.finished_at)}</p>
                  </div>
                </div>
              </div>
              
              {order.observacao && (
                <p className="text-base text-gray-400 mb-4 p-3 bg-gray-800/50 rounded">
                  <span className="font-medium">Observação:</span> {order.observacao}
                </p>
              )}

              {order.laudo_final && (
                <p className="text-base text-green-400 mb-4 p-3 bg-green-900/20 rounded border border-green-700">
                  <span className="font-medium">Laudo Final:</span> {order.laudo_final}
                </p>
              )}
              
              {order.status !== "finalizado" && (
                <div className="flex space-x-3">
                  {order.status === "pendente" && (user.role === "admin" || user.role === "operador_externo" || user.role === "operador_interno") && (
                    <>
                      <Button
                        size="lg"
                        className="btn-merco"
                        onClick={() => startOrderFromList(order)}
                      >
                        Iniciar Produção
                      </Button>
                      <Button
                        size="lg"
                        variant="destructive"
                        onClick={() => deleteOrder(order)}
                      >
                        Excluir
                      </Button>
                    </>
                  )}
                  {order.status === "em_producao" && (
                    <Button
                      size="lg"
                      className="bg-green-600 hover:bg-green-700 text-white"
                      onClick={() => finishOrderFromList(order)}
                    >
                      Finalizar Produção
                    </Button>
                  )}
                </div>
              )}
            </CardContent>
          </Card>
        ))}
      </div>

      {/* Finish Order Dialog */}
      <Dialog open={!!selectedOrder} onOpenChange={() => setSelectedOrder(null)}>
        <DialogContent className="dialog-merco">
          <DialogHeader className="dialog-header">
            <DialogTitle className="dialog-title">
              Finalizar Produção - {selectedOrder?.machine_code}
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-4 p-6 form-merco">
            <div>
              <Label htmlFor="observacao_liberacao">Observação de Liberação</Label>
              <Textarea
                id="observacao_liberacao"
                value={finishData.observacao_liberacao}
                onChange={(e) => setFinishData({...finishData, observacao_liberacao: e.target.value})}
                placeholder="Observações sobre a liberação"
              />
            </div>
            <div>
              <Label htmlFor="laudo_final">Laudo Final *</Label>
              <Textarea
                id="laudo_final"
                value={finishData.laudo_final}
                onChange={(e) => setFinishData({...finishData, laudo_final: e.target.value})}
                placeholder="Laudo final da produção"
                required
              />
            </div>
            <div className="flex space-x-3">
              <Button onClick={handleFinish} className="flex-1 btn-merco" disabled={!finishData.laudo_final}>
                Finalizar Produção
              </Button>
              <Button variant="outline" onClick={() => setSelectedOrder(null)} className="flex-1">
                Cancelar
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>
    </div>
  );
};

// Ordem de Producao Panel
const OrdemProducaoPanel = ({ user }) => {
  const [ordens, setOrdens] = useState([]);
  const [showForm, setShowForm] = useState(false);
  const [artigoSuggestions, setArtigoSuggestions] = useState([]);
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [ordemData, setOrdemData] = useState({
    cliente: "",
    artigo: "",
    cor: "",
    metragem: "",
    data_entrega: "",
    observacao: "",
    engrenagem: "",
    fios: "",
    maquinas: ""
  });
  const suggestionsRef = useRef(null);

  useEffect(() => {
    loadOrdens();
    const interval = setInterval(loadOrdens, 5000);
    return () => clearInterval(interval);
  }, []);

  // Close suggestions when clicking outside
  useEffect(() => {
    const handleClickOutside = (event) => {
      if (suggestionsRef.current && !suggestionsRef.current.contains(event.target)) {
        setShowSuggestions(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const loadOrdens = async () => {
    try {
      const response = await axios.get(`${API}/ordens-producao`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      setOrdens(response.data);
    } catch (error) {
      toast.error("Erro ao carregar ordens de produção");
    }
  };

  // Search artigos for autocomplete
  const searchArtigos = async (term) => {
    if (term.length < 2) {
      setArtigoSuggestions([]);
      setShowSuggestions(false);
      return;
    }
    try {
      const response = await axios.get(`${API}/banco-dados/search?q=${encodeURIComponent(term)}`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      setArtigoSuggestions(response.data);
      setShowSuggestions(response.data.length > 0);
    } catch (error) {
      console.error("Erro ao buscar artigos:", error);
    }
  };

  // Handle artigo input change
  const handleArtigoChange = (e) => {
    const value = e.target.value;
    setOrdemData({ ...ordemData, artigo: value });
    searchArtigos(value);
  };

  // Select artigo from suggestions
  const selectArtigo = (artigo) => {
    setOrdemData({
      ...ordemData,
      artigo: artigo.artigo,
      engrenagem: artigo.engrenagem || "",
      fios: artigo.fios || "",
      maquinas: artigo.maquinas || ""
    });
    setShowSuggestions(false);
    setArtigoSuggestions([]);
  };

  const createOrdem = async () => {
    try {
      await axios.post(`${API}/ordens-producao`, ordemData, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Ordem de produção criada com sucesso!");
      setOrdemData({
        cliente: "",
        artigo: "",
        cor: "",
        metragem: "",
        data_entrega: "",
        observacao: "",
        engrenagem: "",
        fios: "",
        maquinas: ""
      });
      setShowForm(false);
      loadOrdens();
    } catch (error) {
      toast.error("Erro ao criar ordem de produção");
    }
  };

  // Delete ordem de producao
  const deleteOrdem = async (ordemId) => {
    if (window.confirm("Tem certeza que deseja excluir esta ordem de produção?")) {
      try {
        await axios.delete(`${API}/ordens-producao/${ordemId}`, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        });
        toast.success("Ordem de produção excluída com sucesso!");
        loadOrdens();
      } catch (error) {
        toast.error("Erro ao excluir ordem de produção");
      }
    }
  };

  const formatDateTimeBrazil = (utcString) => {
    if (!utcString) return "-";
    try {
      const utcDate = new Date(utcString + (utcString.includes('Z') ? '' : 'Z'));
      return new Intl.DateTimeFormat('pt-BR', {
        timeZone: 'America/Sao_Paulo',
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false
      }).format(utcDate);
    } catch (error) {
      return "-";
    }
  };

  const formatDateBrazil = (dateString) => {
    if (!dateString) return "-";
    try {
      // Parse date string as local date to avoid timezone issues
      const [year, month, day] = dateString.split('T')[0].split('-');
      return `${day}/${month}/${year}`;
    } catch (error) {
      return "-";
    }
  };

  const getStatusBadge = (status) => {
    switch (status) {
      case "pendente": return "bg-yellow-600 text-yellow-100";
      case "em_producao": return "bg-blue-600 text-blue-100";
      case "finalizado": return "bg-green-600 text-green-100";
      default: return "bg-gray-600 text-gray-100";
    }
  };

  const getStatusText = (status) => {
    switch (status) {
      case "pendente": return "Pendente";
      case "em_producao": return "Em Produção";
      case "finalizado": return "Finalizado";
      default: return status;
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <h2 className="text-3xl font-bold text-white">Ordens de Produção</h2>
        <Button onClick={() => setShowForm(!showForm)} className="btn-merco">
          {showForm ? "Cancelar" : "+ Lançar"}
        </Button>
      </div>

      {showForm && (
        <Card className="card-merco">
          <CardHeader className="card-header-merco">
            <CardTitle className="card-title-merco">Nova Ordem de Produção</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4 form-merco">
            <div className="grid grid-cols-2 gap-4">
              <div>
                <Label htmlFor="cliente">Cliente *</Label>
                <Input
                  id="cliente"
                  value={ordemData.cliente}
                  onChange={(e) => setOrdemData({...ordemData, cliente: e.target.value})}
                  placeholder="Nome do cliente"
                  required
                />
              </div>
              <div className="relative" ref={suggestionsRef}>
                <Label htmlFor="artigo">Artigo * (digite para buscar)</Label>
                <Input
                  id="artigo"
                  value={ordemData.artigo}
                  onChange={handleArtigoChange}
                  placeholder="Digite para buscar artigo..."
                  autoComplete="off"
                  required
                />
                {showSuggestions && artigoSuggestions.length > 0 && (
                  <div className="absolute z-50 w-full mt-1 bg-gray-800 border border-gray-600 rounded-md shadow-lg max-h-60 overflow-y-auto">
                    {artigoSuggestions.map((artigo) => (
                      <div
                        key={artigo.id}
                        className="p-3 hover:bg-gray-700 cursor-pointer border-b border-gray-700 last:border-b-0"
                        onClick={() => selectArtigo(artigo)}
                      >
                        <div className="font-medium text-white">{artigo.artigo}</div>
                        <div className="text-xs text-gray-400 mt-1">
                          {artigo.engrenagem && <span>Engr: {artigo.engrenagem} | </span>}
                          {artigo.fios && <span>Fios: {artigo.fios} | </span>}
                          {artigo.maquinas && <span>Máq: {artigo.maquinas}</span>}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
            <div className="grid grid-cols-3 gap-4">
              <div>
                <Label htmlFor="engrenagem">Engrenagem</Label>
                <Input
                  id="engrenagem"
                  value={ordemData.engrenagem}
                  onChange={(e) => setOrdemData({...ordemData, engrenagem: e.target.value})}
                  placeholder="Engrenagem"
                />
              </div>
              <div>
                <Label htmlFor="fios">Fios</Label>
                <Input
                  id="fios"
                  value={ordemData.fios}
                  onChange={(e) => setOrdemData({...ordemData, fios: e.target.value})}
                  placeholder="Quantidade de fios"
                />
              </div>
              <div>
                <Label htmlFor="maquinas">Máquinas</Label>
                <Input
                  id="maquinas"
                  value={ordemData.maquinas}
                  onChange={(e) => setOrdemData({...ordemData, maquinas: e.target.value})}
                  placeholder="Máquinas"
                />
              </div>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <Label htmlFor="cor">Cor *</Label>
                <Input
                  id="cor"
                  value={ordemData.cor}
                  onChange={(e) => setOrdemData({...ordemData, cor: e.target.value})}
                  placeholder="Cor"
                  required
                />
              </div>
              <div>
                <Label htmlFor="metragem">Metragem *</Label>
                <Input
                  id="metragem"
                  value={ordemData.metragem}
                  onChange={(e) => setOrdemData({...ordemData, metragem: e.target.value})}
                  placeholder="Metragem"
                  required
                />
              </div>
            </div>
            <div>
              <Label htmlFor="data_entrega">Data de Entrega *</Label>
              <Input
                id="data_entrega"
                type="date"
                value={ordemData.data_entrega}
                onChange={(e) => setOrdemData({...ordemData, data_entrega: e.target.value})}
                required
              />
            </div>
            <div>
              <Label htmlFor="observacao">Observação</Label>
              <Textarea
                id="observacao"
                value={ordemData.observacao}
                onChange={(e) => setOrdemData({...ordemData, observacao: e.target.value})}
                placeholder="Observações"
              />
            </div>
            <Button 
              onClick={createOrdem} 
              className="w-full btn-merco"
              disabled={!ordemData.cliente || !ordemData.artigo || !ordemData.cor || !ordemData.metragem || !ordemData.data_entrega}
            >
              Criar Ordem
            </Button>
          </CardContent>
        </Card>
      )}

      <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-3">
        {ordens.map((ordem) => (
          <Card key={ordem.id} className="card-merco">
            <CardHeader className="card-header-merco pb-3">
              <div className="flex justify-between items-start">
                <div>
                  <CardTitle className="text-lg text-white font-bold">OS {ordem.numero_os}</CardTitle>
                  <p className="text-sm text-gray-400 mt-1">{ordem.cliente}</p>
                </div>
                <Badge className={getStatusBadge(ordem.status)}>
                  {getStatusText(ordem.status)}
                </Badge>
              </div>
            </CardHeader>
            <CardContent className="space-y-2 text-sm">
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <p className="text-gray-400">Artigo</p>
                  <p className="text-white font-medium">{ordem.artigo}</p>
                </div>
                <div>
                  <p className="text-gray-400">Cor</p>
                  <p className="text-white font-medium">{ordem.cor}</p>
                </div>
              </div>
              {(ordem.engrenagem || ordem.fios || ordem.maquinas) && (
                <div className="grid grid-cols-3 gap-2">
                  {ordem.engrenagem && (
                    <div>
                      <p className="text-gray-400">Engrenagem</p>
                      <p className="text-white font-medium text-xs">{ordem.engrenagem}</p>
                    </div>
                  )}
                  {ordem.fios && (
                    <div>
                      <p className="text-gray-400">Fios</p>
                      <p className="text-white font-medium text-xs">{ordem.fios}</p>
                    </div>
                  )}
                  {ordem.maquinas && (
                    <div>
                      <p className="text-gray-400">Máquinas</p>
                      <p className="text-white font-medium text-xs">{ordem.maquinas}</p>
                    </div>
                  )}
                </div>
              )}
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <p className="text-gray-400">Metragem</p>
                  <p className="text-white font-medium">{ordem.metragem}</p>
                </div>
                <div>
                  <p className="text-gray-400">Entrega</p>
                  <p className="text-white font-medium">{formatDateBrazil(ordem.data_entrega)}</p>
                </div>
              </div>
              {ordem.observacao && (
                <div>
                  <p className="text-gray-400">Observação</p>
                  <p className="text-white text-xs">{ordem.observacao}</p>
                </div>
              )}
              <div className="pt-2 border-t border-gray-700">
                <p className="text-xs text-gray-400">Criado: {formatDateTimeBrazil(ordem.criado_em)}</p>
                {ordem.iniciado_em && (
                  <p className="text-xs text-gray-400">Iniciado: {formatDateTimeBrazil(ordem.iniciado_em)}</p>
                )}
                {ordem.finalizado_em && (
                  <p className="text-xs text-gray-400">Finalizado: {formatDateTimeBrazil(ordem.finalizado_em)}</p>
                )}
                <p className="text-xs text-gray-400">Por: {ordem.criado_por}</p>
              </div>
              {ordem.status === "pendente" && (
                <div className="pt-2">
                  <Button
                    onClick={() => deleteOrdem(ordem.id)}
                    variant="destructive"
                    size="sm"
                    className="w-full"
                    data-testid={`delete-ordem-${ordem.numero_os}`}
                  >
                    Excluir Ordem
                  </Button>
                </div>
              )}
            </CardContent>
          </Card>
        ))}
      </div>
    </div>
  );
};

// Relatorios Panel - Shows only pending ordens
const RelatoriosPanel = ({ user }) => {
  const [ordensPendentes, setOrdensPendentes] = useState([]);
  const [selectedOrdem, setSelectedOrdem] = useState(null);
  const [showEspulaForm, setShowEspulaForm] = useState(false);
  const [machines, setMachines] = useState([]);
  const [machineAllocations, setMachineAllocations] = useState([
    { machine_code: "", machine_id: "", layout_type: "", quantidade: "" }
  ]);
  const [cargasFracoes, setCargasFracoes] = useState([""]);
  const [espulaData, setEspulaData] = useState({
    numero_os: "",
    maquina: "",
    mat_prima: "",
    qtde_fios: "",
    quantidade_metros: "",
    carga: "",
    observacoes: ""
  });

  useEffect(() => {
    loadOrdensPendentes();
    loadMachines();
    const interval = setInterval(loadOrdensPendentes, 5000); // Auto-refresh every 5 seconds
    return () => clearInterval(interval);
  }, []);

  const loadMachines = async () => {
    try {
      const [machines16, machines32] = await Promise.all([
        axios.get(`${API}/machines/16_fusos`, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        }),
        axios.get(`${API}/machines/32_fusos`, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        })
      ]);
      const allMachines = [...machines16.data, ...machines32.data];
      setMachines(allMachines);
    } catch (error) {
      console.error("Erro ao carregar máquinas:", error);
    }
  };

  const addMachineAllocation = () => {
    setMachineAllocations([...machineAllocations, { machine_code: "", machine_id: "", layout_type: "", quantidade: "" }]);
  };

  const removeMachineAllocation = (index) => {
    const newAllocations = machineAllocations.filter((_, i) => i !== index);
    setMachineAllocations(newAllocations);
  };

  const updateMachineAllocation = (index, field, value) => {
    const newAllocations = [...machineAllocations];
    
    if (field === 'machine') {
      const selectedMachine = machines.find(m => m.code === value);
      if (selectedMachine) {
        newAllocations[index] = {
          ...newAllocations[index],
          machine_code: selectedMachine.code,
          machine_id: selectedMachine.id,
          layout_type: selectedMachine.layout_type
        };
      }
    } else if (field === 'quantidade') {
      // Format number
      const formatted = formatNumber(value);
      newAllocations[index].quantidade = formatted;
    }
    
    setMachineAllocations(newAllocations);
  };

  const getTotalQuantidade = () => {
    return machineAllocations.reduce((total, alloc) => {
      const qty = parseInt(alloc.quantidade.replace(/\D/g, '') || '0');
      return total + qty;
    }, 0);
  };

  const addCargaFracao = () => {
    setCargasFracoes([...cargasFracoes, ""]);
  };

  const removeCargaFracao = (index) => {
    if (cargasFracoes.length > 1) {
      const newCargasFracoes = cargasFracoes.filter((_, i) => i !== index);
      setCargasFracoes(newCargasFracoes);
    }
  };

  const updateCargaFracao = (index, value) => {
    const newCargasFracoes = [...cargasFracoes];
    newCargasFracoes[index] = value;
    setCargasFracoes(newCargasFracoes);
  };

  const loadOrdensPendentes = async () => {
    try {
      const response = await axios.get(`${API}/ordens-producao/pendentes`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      // Ordenar por data de entrega (mais próximo primeiro)
      const sortedOrdens = response.data.sort((a, b) => {
        return new Date(a.data_entrega) - new Date(b.data_entrega);
      });
      setOrdensPendentes(sortedOrdens);
    } catch (error) {
      console.error("Erro ao carregar ordens pendentes:", error);
    }
  };

  // Função para formatar número (adiciona pontos de milhar)
  const formatNumber = (value) => {
    // Remove tudo que não é dígito
    const numbers = value.replace(/\D/g, '');
    // Adiciona pontos de milhar
    return numbers.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  };

  const handleOrdemClick = (ordem) => {
    setSelectedOrdem(ordem);
    
    // Se existem dados temporários salvos, carregar eles
    if (ordem.espula_data_temp && Object.keys(ordem.espula_data_temp).length > 0) {
      setEspulaData(ordem.espula_data_temp);
      // Carregar cargas/frações dos dados temporários
      if (ordem.espula_data_temp.cargas_fracoes) {
        setCargasFracoes(ordem.espula_data_temp.cargas_fracoes);
      } else {
        setCargasFracoes([""]);
      }
    } else {
      setEspulaData({
        numero_os: ordem.numero_os,
        maquina: "",
        mat_prima: "",
        qtde_fios: "",
        quantidade_metros: formatNumber(ordem.metragem),
        carga: "",
        observacoes: ordem.observacao || ""
      });
      setCargasFracoes([""]);
    }
    
    // Se existem alocações de máquina temporárias, carregar elas
    if (ordem.dados_temporarios_maquinas && ordem.dados_temporarios_maquinas.length > 0) {
      setMachineAllocations(ordem.dados_temporarios_maquinas);
    } else {
      setMachineAllocations([{ machine_code: "", machine_id: "", layout_type: "", quantidade: "" }]);
    }
    
    setShowEspulaForm(true);
  };

  const saveTempData = async () => {
    try {
      const tempDataPayload = {
        dados_temporarios_maquinas: machineAllocations,
        espula_data: {
          ...espulaData,
          cargas_fracoes: cargasFracoes
        }
      };

      await axios.put(
        `${API}/ordens-producao/${selectedOrdem.id}/salvar-temporarios`,
        tempDataPayload,
        {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        }
      );
      
      toast.success("Dados salvos com sucesso! Você pode continuar depois.");
      setShowEspulaForm(false);
      setSelectedOrdem(null);
      loadOrdensPendentes();
    } catch (error) {
      console.error("Erro ao salvar dados temporários:", error.response?.data);
      toast.error(error.response?.data?.detail || "Erro ao salvar dados temporários");
    }
  };

  const createEspulaFromOrdem = async () => {
    try {
      // Validar alocações de máquinas
      const validAllocations = machineAllocations.filter(a => a.machine_code && a.quantidade);
      
      if (validAllocations.length === 0) {
        toast.error("Selecione pelo menos uma máquina com quantidade");
        return;
      }

      const espulaPayload = {
        ordem_producao_id: selectedOrdem.id,
        numero_os: espulaData.numero_os,
        cliente: selectedOrdem.cliente,
        artigo: selectedOrdem.artigo,
        cor: selectedOrdem.cor,
        maquina: espulaData.maquina,
        mat_prima: espulaData.mat_prima,
        qtde_fios: espulaData.qtde_fios,
        quantidade_metros: espulaData.quantidade_metros,
        carga: espulaData.carga,
        cargas_fracoes: cargasFracoes.filter(cf => cf.trim() !== ""),
        machine_allocations: validAllocations,
        observacoes: espulaData.observacoes,
        data_prevista_entrega: selectedOrdem.data_entrega
      };

      await axios.post(`${API}/espulas`, espulaPayload, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Espulagem criada com sucesso! Ordem movida para produção.");
      setShowEspulaForm(false);
      setSelectedOrdem(null);
      loadOrdensPendentes();
    } catch (error) {
      console.error("Erro ao criar espulagem:", error.response?.data);
      toast.error(error.response?.data?.detail || "Erro ao criar espulagem");
    }
  };

  const formatDateBrazil = (dateString) => {
    if (!dateString) return "-";
    try {
      // Parse date string as local date to avoid timezone issues
      const [year, month, day] = dateString.split('T')[0].split('-');
      return `${day}/${month}/${year}`;
    } catch (error) {
      return "-";
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <h2 className="text-3xl font-bold text-white">Relatórios - Ordens Pendentes</h2>
        <Button onClick={loadOrdensPendentes} variant="outline" className="gap-2">
          <RefreshCw className="h-4 w-4" />
          Atualizar
        </Button>
      </div>

      {ordensPendentes.length === 0 ? (
        <Card className="card-merco">
          <CardContent className="py-12 text-center">
            <p className="text-gray-400 text-lg">Nenhuma ordem pendente no momento</p>
          </CardContent>
        </Card>
      ) : (
        <div className="space-y-3">
          {ordensPendentes.map((ordem) => (
            <Card 
              key={ordem.id} 
              className="card-merco cursor-pointer hover:border-blue-500 transition-colors"
              onClick={() => handleOrdemClick(ordem)}
            >
              <CardContent className="p-4">
                <div className="flex items-center justify-between gap-4 flex-wrap">
                  <div className="flex items-center gap-4 flex-1 min-w-0">
                    <div className="flex-shrink-0">
                      <div className="text-white font-bold text-lg">OS {ordem.numero_os}</div>
                      <Badge className="bg-yellow-600 text-yellow-100 mt-1">Pendente</Badge>
                    </div>
                    <div className="flex-1 grid grid-cols-2 md:grid-cols-4 gap-3 min-w-0">
                      <div className="min-w-0">
                        <p className="text-xs text-gray-400">Cliente</p>
                        <p className="text-white font-medium truncate">{ordem.cliente}</p>
                      </div>
                      <div className="min-w-0">
                        <p className="text-xs text-gray-400">Artigo</p>
                        <p className="text-white font-medium truncate">{ordem.artigo}</p>
                      </div>
                      <div className="min-w-0">
                        <p className="text-xs text-gray-400">Cor</p>
                        <p className="text-white font-medium truncate">{ordem.cor}</p>
                      </div>
                      <div className="min-w-0">
                        <p className="text-xs text-gray-400">Metragem</p>
                        <p className="text-white font-medium">{ordem.metragem}</p>
                      </div>
                    </div>
                  </div>
                  <div className="flex items-center gap-3 flex-shrink-0">
                    <div className="text-right">
                      <p className="text-xs text-gray-400">Data Entrega</p>
                      <p className="text-white font-bold">{formatDateBrazil(ordem.data_entrega)}</p>
                    </div>
                    <Button 
                      size="sm" 
                      className="bg-blue-600 hover:bg-blue-700"
                      onClick={(e) => {
                        e.stopPropagation();
                        handleOrdemClick(ordem);
                      }}
                    >
                      Criar Espulagem
                    </Button>
                  </div>
                </div>
              </CardContent>
            </Card>
          ))}
        </div>
      )}

      {/* Dialog to create Espula from Ordem */}
      <Dialog open={showEspulaForm} onOpenChange={setShowEspulaForm}>
        <DialogContent className="dialog-merco max-w-3xl">
          <DialogHeader className="dialog-header">
            <DialogTitle className="dialog-title">
              Criar Espulagem - OS {selectedOrdem?.numero_os}
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-4 p-6 form-merco max-h-[70vh] overflow-y-auto">
            <div className="grid grid-cols-2 gap-4 p-4 bg-gray-800 rounded">
              <div>
                <p className="text-xs text-gray-400">Cliente</p>
                <p className="text-white font-medium">{selectedOrdem?.cliente}</p>
              </div>
              <div>
                <p className="text-xs text-gray-400">Artigo</p>
                <p className="text-white font-medium">{selectedOrdem?.artigo}</p>
              </div>
              <div>
                <p className="text-xs text-gray-400">Cor</p>
                <p className="text-white font-medium">{selectedOrdem?.cor}</p>
              </div>
              <div>
                <p className="text-xs text-gray-400">Data Entrega</p>
                <p className="text-white font-medium">{formatDateBrazil(selectedOrdem?.data_entrega)}</p>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <div>
                <Label htmlFor="mat_prima">Matéria Prima *</Label>
                <Input
                  id="mat_prima"
                  value={espulaData.mat_prima}
                  onChange={(e) => setEspulaData({...espulaData, mat_prima: e.target.value})}
                  placeholder="Matéria prima"
                  required
                />
              </div>
              <div>
                <Label htmlFor="qtde_fios">Qtde Fios *</Label>
                <Input
                  id="qtde_fios"
                  value={espulaData.qtde_fios}
                  onChange={(e) => setEspulaData({...espulaData, qtde_fios: e.target.value})}
                  placeholder="Quantidade"
                  required
                />
              </div>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <div>
                <Label htmlFor="quantidade_metros">Qtde Metros *</Label>
                <Input
                  id="quantidade_metros"
                  value={espulaData.quantidade_metros}
                  onChange={(e) => {
                    const formatted = formatNumber(e.target.value);
                    setEspulaData({...espulaData, quantidade_metros: formatted});
                  }}
                  placeholder="Ex: 1000 ou 1.000"
                  required
                />
              </div>
              <div>
                <Label htmlFor="carga">Carga *</Label>
                <Input
                  id="carga"
                  value={espulaData.carga}
                  onChange={(e) => setEspulaData({...espulaData, carga: e.target.value})}
                  placeholder="Carga"
                  required
                />
              </div>
            </div>

            <div>
              <div className="flex justify-between items-center mb-2">
                <Label>Cargas e Fração (opcional)</Label>
                <Button size="sm" onClick={addCargaFracao} variant="outline">
                  + Adicionar Carga/Fração
                </Button>
              </div>
              <div className="space-y-2">
                {cargasFracoes.map((carga, index) => (
                  <div key={index} className="flex gap-2 items-center">
                    <Input
                      value={carga}
                      onChange={(e) => updateCargaFracao(index, e.target.value)}
                      placeholder={`Carga/Fração ${index + 1}`}
                      type="text"
                      className="flex-1"
                    />
                    {cargasFracoes.length > 1 && (
                      <Button 
                        size="sm" 
                        variant="ghost" 
                        className="text-red-400 hover:text-red-300"
                        onClick={() => removeCargaFracao(index)}
                      >
                        <Trash2 className="h-4 w-4" />
                      </Button>
                    )}
                  </div>
                ))}
              </div>
            </div>

            <div className="border-t border-gray-700 pt-4">
              <div className="flex justify-between items-center mb-3">
                <Label className="text-lg">Seleção de Máquinas *</Label>
                <Button size="sm" onClick={addMachineAllocation} variant="outline">
                  + Adicionar Máquina
                </Button>
              </div>
              
              {machineAllocations.map((allocation, index) => (
                <div key={index} className="grid grid-cols-3 gap-3 mb-3 p-3 bg-black/40 rounded border border-gray-700">
                  <div className="col-span-2">
                    <Label>Máquina</Label>
                    <Select 
                      value={allocation.machine_code} 
                      onValueChange={(value) => updateMachineAllocation(index, 'machine', value)}
                    >
                      <SelectTrigger>
                        <SelectValue placeholder="Selecione a máquina" />
                      </SelectTrigger>
                      <SelectContent>
                        {machines.map((machine) => (
                          <SelectItem key={machine.id} value={machine.code}>
                            {machine.code} - {machine.layout_type === '16_fusos' ? '16 Fusos' : '32 Fusos'}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="relative">
                    <Label>Quantidade</Label>
                    <Input
                      value={allocation.quantidade}
                      onChange={(e) => updateMachineAllocation(index, 'quantidade', e.target.value)}
                      placeholder="Ex: 1.000"
                    />
                    {machineAllocations.length > 1 && (
                      <Button 
                        size="sm" 
                        variant="ghost" 
                        className="absolute top-0 right-0 text-red-400"
                        onClick={() => removeMachineAllocation(index)}
                      >
                        <Trash2 className="h-4 w-4" />
                      </Button>
                    )}
                  </div>
                </div>
              ))}
              
              <div className="flex justify-between items-center p-3 bg-blue-900/30 rounded mt-2">
                <span className="text-white font-medium">Total:</span>
                <span className="text-white text-lg font-bold">{getTotalQuantidade().toLocaleString('pt-BR')}</span>
              </div>
            </div>

            <div>
              <Label htmlFor="observacoes">Observações</Label>
              <Textarea
                id="observacoes"
                value={espulaData.observacoes}
                onChange={(e) => setEspulaData({...espulaData, observacoes: e.target.value})}
                placeholder="Observações adicionais"
              />
            </div>

            <div className="flex space-x-3">
              <Button 
                onClick={createEspulaFromOrdem} 
                className="flex-1 btn-merco"
                disabled={!espulaData.mat_prima || !espulaData.qtde_fios || !espulaData.quantidade_metros || !espulaData.carga || machineAllocations.filter(a => a.machine_code && a.quantidade).length === 0}
              >
                Criar Espulagens
              </Button>
              <Button 
                onClick={saveTempData} 
                variant="outline"
                className="flex-1 bg-blue-600 hover:bg-blue-700 text-white"
              >
                Salvar
              </Button>
              <Button variant="outline" onClick={() => setShowEspulaForm(false)} className="flex-1">
                Cancelar
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>
    </div>
  );
};

const EspulasPanel = ({ espulas, user, onEspulaUpdate }) => {
  const [showForm, setShowForm] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [cargasFracoes, setCargasFracoes] = useState([""]);
  const [editingMachines, setEditingMachines] = useState(null);
  const [machineAllocations, setMachineAllocations] = useState([]);
  const [machines, setMachines] = useState([]);
  const [espulaData, setEspulaData] = useState({
    numero_os: "",
    maquina: "",
    mat_prima: "",
    qtde_fios: "",
    cliente: "",
    artigo: "",
    cor: "",
    quantidade_metros: "",
    carga: "",
    observacoes: "",
    data_prevista_entrega: ""
  });

  // Função para formatar número (adiciona pontos de milhar)
  const formatNumber = (value) => {
    // Remove tudo que não é dígito
    const numbers = value.replace(/\D/g, '');
    // Adiciona pontos de milhar
    return numbers.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  };

  // Função para lidar com mudança de metros
  const handleMetrosChange = (value) => {
    const formatted = formatNumber(value);
    setEspulaData({...espulaData, quantidade_metros: formatted});
  };

  // Funções para gerenciar cargas e frações
  const addCargaFracao = () => {
    setCargasFracoes([...cargasFracoes, ""]);
  };

  const removeCargaFracao = (index) => {
    if (cargasFracoes.length > 1) {
      const newCargasFracoes = cargasFracoes.filter((_, i) => i !== index);
      setCargasFracoes(newCargasFracoes);
    }
  };

  const updateCargaFracao = (index, value) => {
    const newCargasFracoes = [...cargasFracoes];
    newCargasFracoes[index] = value;
    setCargasFracoes(newCargasFracoes);
  };

  // Função para buscar próximo número de OS
  const getNextOSNumber = async () => {
    try {
      const response = await axios.get(`${API}/ordens-producao/next-number`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      return response.data.numero_os;
    } catch (error) {
      console.error("Erro ao buscar próximo número:", error);
      return "";
    }
  };

  // Carregar máquinas
  useEffect(() => {
    loadMachines();
  }, []);

  // Carregar alocações quando editingMachines muda
  useEffect(() => {
    if (editingMachines) {
      const allocations = editingMachines.machine_allocations;
      if (allocations && allocations.length > 0) {
        const copiedAllocations = allocations.map(alloc => ({
          machine_code: alloc.machine_code || "",
          machine_id: alloc.machine_id || "",
          layout_type: alloc.layout_type || "",
          quantidade: alloc.quantidade || ""
        }));
        setMachineAllocations(copiedAllocations);
      } else {
        setMachineAllocations([{ machine_code: "", machine_id: "", layout_type: "", quantidade: "" }]);
      }
    }
  }, [editingMachines]);

  const loadMachines = async () => {
    try {
      const response = await axios.get(`${API}/machines`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      setMachines(response.data);
    } catch (error) {
      console.error("Erro ao carregar máquinas:", error);
    }
  };

  // Funções para gerenciar alocações de máquinas
  const addMachineAllocation = () => {
    setMachineAllocations([...machineAllocations, { machine_code: "", machine_id: "", layout_type: "", quantidade: "" }]);
  };

  const removeMachineAllocation = (index) => {
    if (machineAllocations.length > 1) {
      const newAllocations = machineAllocations.filter((_, i) => i !== index);
      setMachineAllocations(newAllocations);
    }
  };

  const updateMachineAllocation = (index, field, value) => {
    const newAllocations = [...machineAllocations];
    
    if (field === 'machine') {
      const selectedMachine = machines.find(m => m.code === value);
      if (selectedMachine) {
        newAllocations[index].machine_code = selectedMachine.code;
        newAllocations[index].machine_id = selectedMachine.id;
        newAllocations[index].layout_type = selectedMachine.layout_type;
      }
    } else if (field === 'quantidade') {
      newAllocations[index].quantidade = formatNumber(value);
    }
    
    setMachineAllocations(newAllocations);
  };

  // Abrir dialog para editar máquinas de uma espulagem
  const openEditMachines = async (espula) => {
    // Garantir que máquinas estão carregadas
    await loadMachines();
    // Setar editingMachines vai disparar o useEffect que carrega as alocações
    setEditingMachines(espula);
  };

  // Salvar as alterações de máquinas
  const saveMachineAllocations = async () => {
    try {
      await axios.put(
        `${API}/espulas/${editingMachines.id}/machine-allocations`,
        { machine_allocations: machineAllocations },
        { headers: { Authorization: `Bearer ${localStorage.getItem("token")}` } }
      );
      
      toast.success("Máquinas atualizadas com sucesso!");
      setEditingMachines(null);
      onEspulaUpdate();
    } catch (error) {
      console.error("Erro ao atualizar máquinas:", error);
      toast.error("Erro ao atualizar máquinas");
    }
  };

  // Abrir formulário com próximo número de OS
  const openForm = async () => {
    const nextNumber = await getNextOSNumber();
    setEspulaData({
      numero_os: nextNumber,
      maquina: "",
      mat_prima: "",
      qtde_fios: "",
      cliente: "",
      artigo: "",
      cor: "",
      quantidade_metros: "",
      carga: "",
      observacoes: "",
      data_prevista_entrega: ""
    });
    setCargasFracoes([""]);
    setShowForm(true);
  };

  // CORRIGIR FORMATAÇÃO DE HORÁRIO - CONVERSÃO CORRETA UTC PARA BRASÍLIA  
  const formatDateTimeBrazil = (utcString) => {
    if (!utcString) return "-";
    try {
      // Parse UTC string e converte para horário de Brasília
      const utcDate = new Date(utcString + (utcString.includes('Z') ? '' : 'Z'));
      
      return new Intl.DateTimeFormat('pt-BR', {
        timeZone: 'America/Sao_Paulo',
        day: '2-digit',
        month: '2-digit', 
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false
      }).format(utcDate);
    } catch (error) {
      console.error('Error formatting date:', error);
      return "-";
    }
  };

  const formatDateBrazil = (utcString) => {
    if (!utcString) return "-";
    try {
      // Parse UTC string e converte para horário de Brasília
      const utcDate = new Date(utcString + (utcString.includes('Z') ? '' : 'Z'));
      
      return new Intl.DateTimeFormat('pt-BR', {
        timeZone: 'America/Sao_Paulo',
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      }).format(utcDate);
    } catch (error) {
      console.error('Error formatting date:', error);
      return "-";
    }
  };

  const createEspula = async () => {
    try {
      const espulaPayload = {
        ...espulaData,
        cargas_fracoes: cargasFracoes.filter(cf => cf.trim() !== "")
      };
      
      await axios.post(`${API}/espulas`, espulaPayload, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Espulagem lançada com sucesso!");
      setEspulaData({
        numero_os: "",
        maquina: "",
        mat_prima: "",
        qtde_fios: "",
        cliente: "",
        artigo: "",
        cor: "",
        quantidade_metros: "",
        carga: "",
        observacoes: "",
        data_prevista_entrega: ""
      });
      setCargasFracoes([""]);
      setShowForm(false);
      onEspulaUpdate();
    } catch (error) {
      toast.error("Erro ao lançar espulagem");
    }
  };

  const updateEspulaStatus = async (espulaId, newStatus) => {
    try {
      await axios.put(`${API}/espulas/${espulaId}`, { status: newStatus }, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Status atualizado com sucesso!");
      onEspulaUpdate();
    } catch (error) {
      toast.error("Erro ao atualizar status");
    }
  };

  const finalizeEspulaWithMachines = async (espulaId) => {
    try {
      const response = await axios.post(`${API}/espulas/${espulaId}/finalize-with-machines`, {}, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success(`Espulagem finalizada! ${response.data.orders_created} pedidos criados nas máquinas.`);
      onEspulaUpdate();
    } catch (error) {
      toast.error(error.response?.data?.detail || "Erro ao finalizar espulagem");
    }
  };

  const exportEspulasReport = async () => {
    try {
      const XLSXStyle = require('xlsx-js-style');
      
      const pendenteEspulas = espulas.filter(e => e.status === "pendente");
      
      if (pendenteEspulas.length === 0) {
        toast.warning("Nenhuma espulagem pendente para exportar");
        return;
      }
      
      // Buscar dados dos artigos do banco de dados para preencher informações
      let artigosData = [];
      try {
        const response = await axios.get(`${API}/banco-dados`, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        });
        artigosData = response.data;
      } catch (err) {
        console.error("Erro ao carregar artigos:", err);
      }
      
      const wb = XLSXStyle.utils.book_new();
      
      // Descobrir o número máximo de cargas em todas as espulagens
      let maxCargas = 0;
      pendenteEspulas.forEach(espula => {
        if (espula.cargas_fracoes && espula.cargas_fracoes.length > maxCargas) {
          maxCargas = espula.cargas_fracoes.length;
        }
      });
      maxCargas = Math.max(5, Math.min(maxCargas, 20)); // Mínimo 5, máximo 20
      
      const aoa = [];
      
      // Linha 1: Título
      const today = new Date().toLocaleDateString('pt-BR');
      aoa.push([`PEDIDO DE ESPULAGEM ${today}`]);
      
      // Linha 2: Cabeçalhos DINÂMICOS
      const headers = [
        'Artigo', 'MP', 'Maq.', 'ENGRENAGEM', 'ENCHIMENTO', 'CICLOS',
        'Cor 1', 'Cor 2', 'Cor 3', 'Cor 4', 'Fios'
      ];
      
      // Adicionar colunas de Carga dinamicamente
      for (let i = 1; i <= maxCargas; i++) {
        headers.push(`Carga ${i}`);
      }
      headers.push('Carga Total'); // SEMPRE depois da última carga
      
      aoa.push(headers);
      
      // UMA LINHA por espulagem (SEM linhas em branco)
      pendenteEspulas.forEach(espula => {
        // Buscar dados do artigo no banco de dados
        const artigoInfo = artigosData.find(a => 
          a.artigo && espula.artigo && 
          a.artigo.toLowerCase() === espula.artigo.toLowerCase()
        );
        
        // Montar lista de máquinas (prioridade: alocadas > espulagem > banco de dados)
        let maquinasList = '';
        if (espula.machine_allocations && espula.machine_allocations.length > 0) {
          maquinasList = espula.machine_allocations
            .map(alloc => alloc.machine_code)
            .join(', ');
        } else if (espula.maquina) {
          maquinasList = espula.maquina;
        } else if (artigoInfo?.maquinas) {
          // Usar máquinas do banco de dados se não houver máquinas na espulagem
          maquinasList = artigoInfo.maquinas;
        }
        
        // Preencher engrenagem, enchimento (carga) e ciclos com dados reais
        const engrenagem = artigoInfo?.engrenagem || '';
        const enchimento = artigoInfo?.carga || ''; // carga = enchimento
        const ciclos = artigoInfo?.ciclos || '';
        
        // Preencher cores dinamicamente (Cor 1, Cor 2, Cor 3, Cor 4)
        const cor1 = espula.cor || '';
        const cor2 = '';
        const cor3 = '';
        const cor4 = '';
        
        // Calcular Carga Total como SOMA de todas as cargas preenchidas
        let cargaTotal = 0;
        if (espula.cargas_fracoes && espula.cargas_fracoes.length > 0) {
          espula.cargas_fracoes.forEach(carga => {
            const valor = parseFloat(carga);
            if (!isNaN(valor)) {
              cargaTotal += valor;
            }
          });
        }
        
        // Montar linha única
        const row = [
          espula.artigo || '',
          espula.mat_prima || '',
          maquinasList,
          engrenagem, // ENGRENAGEM - dados reais do banco
          enchimento, // ENCHIMENTO - dados reais do banco (carga)
          ciclos, // CICLOS - dados reais do banco
          cor1, // Cor 1 - preenchida
          cor2, // Cor 2
          cor3, // Cor 3
          cor4, // Cor 4
          espula.qtde_fios || artigoInfo?.fios || '',
        ];
        
        // Adicionar TODAS as cargas (até maxCargas)
        for (let i = 0; i < maxCargas; i++) {
          if (espula.cargas_fracoes && espula.cargas_fracoes[i]) {
            row.push(espula.cargas_fracoes[i]);
          } else {
            row.push(''); // Vazio se não tiver
          }
        }
        
        // Carga Total CALCULADA - sempre depois da última carga
        row.push(cargaTotal > 0 ? cargaTotal.toString() : espula.carga || '');
        
        aoa.push(row);
      });
      
      // Criar worksheet
      const ws = XLSXStyle.utils.aoa_to_sheet(aoa);
      
      // Estilos
      const titleStyle = {
        font: { bold: true, sz: 14, color: { rgb: "000000" } },
        alignment: { horizontal: "center", vertical: "center" },
        fill: { fgColor: { rgb: "FFFFFF" } },
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thin", color: { rgb: "000000" } },
          left: { style: "thin", color: { rgb: "000000" } },
          right: { style: "thin", color: { rgb: "000000" } }
        }
      };
      
      const headerStyle = {
        font: { bold: true, sz: 11, color: { rgb: "000000" } },
        alignment: { horizontal: "center", vertical: "center" },
        fill: { fgColor: { rgb: "D3D3D3" } },
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thin", color: { rgb: "000000" } },
          left: { style: "thin", color: { rgb: "000000" } },
          right: { style: "thin", color: { rgb: "000000" } }
        }
      };
      
      const dataStyle = {
        font: { sz: 10, color: { rgb: "000000" } },
        alignment: { horizontal: "left", vertical: "center" },
        fill: { fgColor: { rgb: "FFFFFF" } },
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thin", color: { rgb: "000000" } },
          left: { style: "thin", color: { rgb: "000000" } },
          right: { style: "thin", color: { rgb: "000000" } }
        }
      };
      
      // Aplicar estilo ao título
      ws['A1'].s = titleStyle;
      
      // Número de colunas
      const numCols = headers.length;
      
      // Mesclar título
      if (!ws['!merges']) ws['!merges'] = [];
      ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: numCols - 1 } });
      
      // Gerar letras das colunas
      const cols = [];
      for (let i = 0; i < numCols; i++) {
        if (i < 26) {
          cols.push(String.fromCharCode(65 + i));
        } else {
          const first = String.fromCharCode(65 + Math.floor(i / 26) - 1);
          const second = String.fromCharCode(65 + (i % 26));
          cols.push(first + second);
        }
      }
      
      // Aplicar estilos aos cabeçalhos
      cols.forEach(col => {
        if (ws[`${col}2`]) ws[`${col}2`].s = headerStyle;
      });
      
      // Aplicar estilos aos dados
      const totalRows = aoa.length;
      for (let row = 3; row <= totalRows; row++) {
        cols.forEach((col, idx) => {
          const cell = `${col}${row}`;
          if (ws[cell]) {
            ws[cell].s = { ...dataStyle };
            // Negrito: Artigo (A), MP (B), Maq (C)
            if (idx <= 2 && ws[cell].v) {
              ws[cell].s.font = { ...dataStyle.font, bold: true };
            }
          }
        });
      }
      
      // Larguras das colunas
      ws['!cols'] = cols.map(() => ({ wch: 12 }));
      
      XLSXStyle.utils.book_append_sheet(wb, ws, "PEDIDO DE ESPULAGEM");
      XLSXStyle.writeFile(wb, `pedido_espulagem_${new Date().toISOString().split('T')[0]}.xlsx`);
      toast.success("Relatório exportado com sucesso!");
      
    } catch (error) {
      console.error("Erro ao exportar:", error);
      toast.error("Erro ao exportar relatório");
    }
  };

  const getStatusBadge = (status) => {
    switch (status) {
      case "pendente": return "bg-yellow-600 text-yellow-100";
      case "em_producao_aguardando": return "bg-red-600 text-red-100";
      case "producao": return "bg-red-700 text-red-100";
      case "finalizado": return "bg-green-600 text-green-100";
      default: return "bg-gray-600 text-gray-100";
    }
  };

  const getStatusText = (status) => {
    switch (status) {
      case "pendente": return "Pendente";
      case "em_producao_aguardando": return "Em Produção (Aguardando)";
      case "producao": return "Produção";
      case "finalizado": return "Finalizado";
      default: return status;
    }
  };

  // Sort espulas by priority (delivery date) for active view
  const sortedEspulas = [...espulas].sort((a, b) => {
    if (a.status === "em_producao_aguardando" || a.status === "producao") {
      if (b.status === "em_producao_aguardando" || b.status === "producao") {
        return new Date(a.data_prevista_entrega) - new Date(b.data_prevista_entrega);
      }
      return -1;
    }
    if (b.status === "em_producao_aguardando" || b.status === "producao") {
      return 1;
    }
    return new Date(a.data_prevista_entrega) - new Date(b.data_prevista_entrega);
  });

  // Sort history by finalized date (most recent first)
  const sortedHistory = [...espulas].sort((a, b) => {
    // Se ambos têm finalizado_em, ordenar por data de finalização (mais recente primeiro)
    if (a.finalizado_em && b.finalizado_em) {
      return new Date(b.finalizado_em) - new Date(a.finalizado_em);
    }
    // Se apenas um tem finalizado_em, o que tem vem primeiro
    if (a.finalizado_em) return -1;
    if (b.finalizado_em) return 1;
    // Se nenhum tem finalizado_em, ordenar por data de criação (mais recente primeiro)
    return new Date(b.created_at) - new Date(a.created_at);
  });

  const activeEspulas = sortedEspulas.filter(e => e.status !== "finalizado");
  const finalizedEspulas = sortedHistory.filter(e => e.status === "finalizado");

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <h2 className="text-3xl font-bold text-white">Gerenciamento de Espulagem</h2>
        <div className="flex space-x-3">
          <Button 
            onClick={() => setShowHistory(!showHistory)} 
            className="bg-blue-600 hover:bg-blue-700 text-white"
          >
            {showHistory ? "Ocultar Histórico Espulagem" : "Ver Histórico Espulagem"}
          </Button>
          <Button onClick={exportEspulasReport} className="bg-green-600 hover:bg-green-700 text-white">
            <Download className="h-4 w-4 mr-2" />
            Exportar Relatório
          </Button>
          <Button onClick={() => showForm ? setShowForm(false) : openForm()} className="btn-merco-large">
            {showForm ? "Cancelar" : "+ LANÇAR"}
          </Button>
        </div>
      </div>

      {showForm && (
        <Card className="card-merco-large">
          <CardHeader>
            <CardTitle className="text-white text-xl">Nova Ordem de Serviço - Espulagem</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4 form-merco">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div>
                <Label htmlFor="numero_os">OS (Automático) *</Label>
                <Input
                  id="numero_os"
                  value={espulaData.numero_os}
                  readOnly
                  disabled
                  className="bg-gray-800 cursor-not-allowed"
                  placeholder="Número gerado automaticamente"
                />
              </div>
              <div>
                <Label htmlFor="maquina">Máquina *</Label>
                <Input
                  id="maquina"
                  value={espulaData.maquina}
                  onChange={(e) => setEspulaData({...espulaData, maquina: e.target.value})}
                  placeholder="Ex: CD1, F2"
                  required
                />
              </div>
              <div>
                <Label htmlFor="mat_prima">Matéria Prima *</Label>
                <Input
                  id="mat_prima"
                  value={espulaData.mat_prima}
                  onChange={(e) => setEspulaData({...espulaData, mat_prima: e.target.value})}
                  placeholder="Matéria prima"
                  required
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div>
                <Label htmlFor="cliente">Cliente *</Label>
                <Input
                  id="cliente"
                  value={espulaData.cliente}
                  onChange={(e) => setEspulaData({...espulaData, cliente: e.target.value})}
                  placeholder="Nome do cliente"
                  required
                />
              </div>
              <div>
                <Label htmlFor="artigo">Artigo *</Label>
                <Input
                  id="artigo"
                  value={espulaData.artigo}
                  onChange={(e) => setEspulaData({...espulaData, artigo: e.target.value})}
                  placeholder="Artigo"
                  required
                />
              </div>
              <div>
                <Label htmlFor="cor">Cor *</Label>
                <Input
                  id="cor"
                  value={espulaData.cor}
                  onChange={(e) => setEspulaData({...espulaData, cor: e.target.value})}
                  placeholder="Cor"
                  required
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div>
                <Label htmlFor="qtde_fios">Qtde Fios *</Label>
                <Input
                  id="qtde_fios"
                  value={espulaData.qtde_fios}
                  onChange={(e) => setEspulaData({...espulaData, qtde_fios: e.target.value})}
                  placeholder="Quantidade de fios"
                  required
                />
              </div>
              <div>
                <Label htmlFor="quantidade_metros">Qtde Metros *</Label>
                <Input
                  id="quantidade_metros"
                  value={espulaData.quantidade_metros}
                  onChange={(e) => handleMetrosChange(e.target.value)}
                  placeholder="Ex: 1000 ou 1.000"
                  required
                />
              </div>
              <div>
                <Label htmlFor="data_prevista_entrega">Data Entrega *</Label>
                <Input
                  id="data_prevista_entrega"
                  type="date"
                  value={espulaData.data_prevista_entrega}
                  onChange={(e) => setEspulaData({...espulaData, data_prevista_entrega: e.target.value})}
                  required
                />
              </div>
            </div>

            <div>
              <Label htmlFor="carga">Carga *</Label>
              <Input
                id="carga"
                value={espulaData.carga}
                onChange={(e) => setEspulaData({...espulaData, carga: e.target.value})}
                placeholder="Ex: A1, B2, C123, etc"
                required
              />
            </div>

            <div>
              <div className="flex justify-between items-center mb-2">
                <Label>Cargas e Fração (opcional)</Label>
                <Button size="sm" onClick={addCargaFracao} variant="outline">
                  + Adicionar Carga/Fração
                </Button>
              </div>
              <div className="space-y-2">
                {cargasFracoes.map((carga, index) => (
                  <div key={index} className="flex gap-2 items-center">
                    <Input
                      value={carga}
                      onChange={(e) => updateCargaFracao(index, e.target.value)}
                      placeholder={`Carga/Fração ${index + 1}`}
                      type="text"
                      className="flex-1"
                    />
                    {cargasFracoes.length > 1 && (
                      <Button 
                        size="sm" 
                        variant="ghost" 
                        className="text-red-400 hover:text-red-300"
                        onClick={() => removeCargaFracao(index)}
                      >
                        <Trash2 className="h-4 w-4" />
                      </Button>
                    )}
                  </div>
                ))}
              </div>
            </div>

            <div>
              <Label htmlFor="observacoes">Observações</Label>
              <Textarea
                id="observacoes"
                value={espulaData.observacoes}
                onChange={(e) => setEspulaData({...espulaData, observacoes: e.target.value})}
                placeholder="Observações adicionais"
              />
            </div>
            <Button 
              onClick={createEspula} 
              className="w-full btn-merco-large"
              disabled={!espulaData.maquina || !espulaData.mat_prima || !espulaData.cliente || !espulaData.artigo || !espulaData.cor || !espulaData.qtde_fios || !espulaData.quantidade_metros || !espulaData.carga || !espulaData.data_prevista_entrega}
            >
              + LANÇAR ESPULAGEM
            </Button>
          </CardContent>
        </Card>
      )}

      {/* Active Espulas */}
      {!showHistory && (
        <div className="grid gap-6">
          <h3 className="text-2xl font-bold text-white">Espulagem Ativa</h3>
          {activeEspulas.map((espula) => (
            <Card key={espula.id} className="card-merco-large">
              <CardContent className="p-6">
                <div className="flex justify-between items-start mb-4">
                  <div>
                    <h3 className="font-bold text-white text-xl">
                      {espula.numero_os ? `OS ${espula.numero_os}` : `OS #${espula.id.slice(-8)}`} - {espula.cliente}
                    </h3>
                    <p className="text-gray-400 text-lg">Artigo: <span className="text-white">{espula.artigo}</span></p>
                  </div>
                  <Badge className={`${getStatusBadge(espula.status)} font-semibold text-sm`}>
                    {getStatusText(espula.status)}
                  </Badge>
                </div>
                
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-base mb-4">
                  {espula.maquina && (
                    <div>
                      <span className="font-medium text-gray-400">Máquina:</span>
                      <p className="text-white">{espula.maquina}</p>
                    </div>
                  )}
                  <div>
                    <span className="font-medium text-gray-400">Cor:</span>
                    <p className="text-white">{espula.cor}</p>
                  </div>
                  {espula.mat_prima && (
                    <div>
                      <span className="font-medium text-gray-400">Mat. Prima:</span>
                      <p className="text-white">{espula.mat_prima}</p>
                    </div>
                  )}
                  {espula.qtde_fios && (
                    <div>
                      <span className="font-medium text-gray-400">Qtde Fios:</span>
                      <p className="text-white">{espula.qtde_fios}</p>
                    </div>
                  )}
                  <div>
                    <span className="font-medium text-gray-400">Qtde Metros:</span>
                    <p className="text-white">{espula.quantidade_metros}</p>
                  </div>
                  <div>
                    <span className="font-medium text-gray-400">Carga:</span>
                    <p className="text-white">{espula.carga}</p>
                  </div>
                  <div>
                    <span className="font-medium text-gray-400">Data Entrega:</span>
                    <p className="text-white">{formatDateBrazil(espula.data_prevista_entrega)}</p>
                  </div>
                </div>

                {/* Cargas e Fração */}
                {(espula.cargas_fracoes && espula.cargas_fracoes.length > 0 && espula.cargas_fracoes.some(cf => cf)) && (
                  <div className="mb-4 p-3 bg-gray-800/50 rounded">
                    <span className="font-medium text-gray-400">Cargas e Fração:</span>
                    <div className="flex gap-3 mt-2 flex-wrap">
                      {espula.cargas_fracoes.map((carga, index) => (
                        carga && <span key={index} className="text-white px-3 py-1 bg-gray-700 rounded">{carga}</span>
                      ))}
                    </div>
                  </div>
                )}

                {/* Time history */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-base mb-4 p-4 bg-gray-900/50 rounded-lg">
                  <div className="flex items-center space-x-3">
                    <Clock className="h-5 w-5 text-blue-400" />
                    <div>
                      <span className="font-medium text-gray-400">Lançado:</span>
                      <p className="text-white">{formatDateTimeBrazil(espula.created_at)}</p>
                    </div>
                  </div>
                  {espula.iniciado_em && (
                    <div className="flex items-center space-x-3">
                      <Clock className="h-5 w-5 text-yellow-400" />
                      <div>
                        <span className="font-medium text-gray-400">Iniciado:</span>
                        <p className="text-white">{formatDateTimeBrazil(espula.iniciado_em)}</p>
                      </div>
                    </div>
                  )}
                  {espula.finalizado_em && (
                    <div className="flex items-center space-x-3">
                      <Clock className="h-5 w-5 text-green-400" />
                      <div>
                        <span className="font-medium text-gray-400">Finalizado:</span>
                        <p className="text-white">{formatDateTimeBrazil(espula.finalizado_em)}</p>
                      </div>
                    </div>
                  )}
                </div>

                {espula.observacoes && (
                  <p className="text-base text-gray-400 mb-4 p-3 bg-gray-800/50 rounded">
                    <span className="font-medium">Observações:</span> {espula.observacoes}
                  </p>
                )}

                {/* Máquinas Alocadas */}
                {espula.machine_allocations && espula.machine_allocations.length > 0 && (
                  <div className="mb-4 p-4 bg-blue-900/20 rounded border border-blue-700">
                    <div className="flex justify-between items-center mb-3">
                      <h4 className="font-bold text-white text-lg flex items-center">
                        <Factory className="h-5 w-5 mr-2" />
                        Máquinas Alocadas ({espula.machine_allocations.length})
                      </h4>
                      <Button 
                        size="sm" 
                        onClick={() => openEditMachines(espula)}
                        className="bg-blue-600 hover:bg-blue-700 text-white"
                      >
                        Editar Máquinas
                      </Button>
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                      {espula.machine_allocations.map((allocation, index) => (
                        <div key={index} className="p-3 bg-black/40 rounded border border-gray-700">
                          <div className="flex justify-between items-start mb-2">
                            <span className="font-bold text-white text-lg">{allocation.machine_code}</span>
                            <Badge className="bg-gray-700 text-gray-200 text-xs">
                              {allocation.layout_type === '16_fusos' ? '16 Fusos' : '32 Fusos'}
                            </Badge>
                          </div>
                          <div className="text-sm">
                            <span className="text-gray-400">Quantidade:</span>
                            <span className="text-white font-bold ml-2">{allocation.quantidade}</span>
                          </div>
                        </div>
                      ))}
                    </div>
                    <div className="mt-3 pt-3 border-t border-blue-700/50 flex justify-between items-center">
                      <span className="text-gray-400">Total Geral:</span>
                      <span className="text-white text-xl font-bold">
                        {espula.machine_allocations.reduce((sum, a) => sum + parseInt(a.quantidade.replace(/\D/g, '') || 0), 0).toLocaleString('pt-BR')}
                      </span>
                    </div>
                  </div>
                )}

                {espula.status !== "finalizado" && (
                  <div className="flex space-x-3">
                    {espula.status === "pendente" && (
                      <Button
                        className="bg-red-600 hover:bg-red-700 text-white"
                        onClick={() => updateEspulaStatus(espula.id, "em_producao_aguardando")}
                      >
                        Colocar em Produção
                      </Button>
                    )}
                    {espula.status === "em_producao_aguardando" && (
                      <Button
                        className="bg-red-700 hover:bg-red-800 text-white"
                        onClick={() => updateEspulaStatus(espula.id, "producao")}
                      >
                        Iniciar Produção
                      </Button>
                    )}
                    {espula.status === "producao" && (
                      <>
                        {espula.machine_allocations && espula.machine_allocations.length > 0 ? (
                          <Button
                            className="bg-green-600 hover:bg-green-700 text-white flex items-center gap-2"
                            onClick={() => finalizeEspulaWithMachines(espula.id)}
                          >
                            <Factory className="h-4 w-4" />
                            Finalizar e Enviar para Máquinas
                          </Button>
                        ) : (
                          <Button
                            className="bg-green-600 hover:bg-green-700 text-white"
                            onClick={() => updateEspulaStatus(espula.id, "finalizado")}
                          >
                            Finalizar
                          </Button>
                        )}
                      </>
                    )}
                  </div>
                )}
              </CardContent>
            </Card>
          ))}
          
          {activeEspulas.length === 0 && (
            <Card className="card-merco-large">
              <CardContent className="p-6 text-center">
                <p className="text-gray-400 text-lg">Nenhuma espulagem ativa. Clique em "+ LANÇAR" para criar a primeira.</p>
              </CardContent>
            </Card>
          )}
        </div>
      )}

      {/* History - All Espulas */}
      {showHistory && (
        <div className="grid gap-6">
          <h3 className="text-2xl font-bold text-white">Histórico - Toda Espulagem</h3>
          {sortedHistory.map((espula) => (
            <Card key={espula.id} className="card-merco-large opacity-75">
              <CardContent className="p-6">
                <div className="flex justify-between items-start mb-4">
                  <div>
                    <h3 className="font-bold text-white text-xl">
                      {espula.numero_os ? `OS ${espula.numero_os}` : `OS #${espula.id.slice(-8)}`} - {espula.cliente}
                    </h3>
                    <p className="text-gray-400 text-lg">Artigo: <span className="text-white">{espula.artigo}</span></p>
                  </div>
                  <Badge className={`${getStatusBadge(espula.status)} font-semibold text-sm`}>
                    {getStatusText(espula.status)}
                  </Badge>
                </div>
                
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-base mb-4">
                  <div>
                    <span className="font-medium text-gray-400">Cor:</span>
                    <p className="text-white">{espula.cor}</p>
                  </div>
                  <div>
                    <span className="font-medium text-gray-400">Quantidade (m):</span>
                    <p className="text-white">{espula.quantidade_metros}</p>
                  </div>
                  <div>
                    <span className="font-medium text-gray-400">Carga:</span>
                    <p className="text-white">{espula.carga}</p>
                  </div>
                  <div>
                    <span className="font-medium text-gray-400">Criado por:</span>
                    <p className="text-white">{espula.created_by}</p>
                  </div>
                </div>

                {/* Complete time history */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-base mb-4 p-4 bg-gray-900/50 rounded-lg">
                  <div className="flex items-center space-x-3">
                    <Clock className="h-5 w-5 text-blue-400" />
                    <div>
                      <span className="font-medium text-gray-400">Lançado:</span>
                      <p className="text-white">{formatDateTimeBrazil(espula.created_at)}</p>
                    </div>
                  </div>
                  <div className="flex items-center space-x-3">
                    <Clock className="h-5 w-5 text-yellow-400" />
                    <div>
                      <span className="font-medium text-gray-400">Iniciado:</span>
                      <p className="text-white">{formatDateTimeBrazil(espula.iniciado_em)}</p>
                    </div>
                  </div>
                  <div className="flex items-center space-x-3">
                    <Clock className="h-5 w-5 text-green-400" />
                    <div>
                      <span className="font-medium text-gray-400">Finalizado:</span>
                      <p className="text-white">{formatDateTimeBrazil(espula.finalizado_em)}</p>
                    </div>
                  </div>
                </div>

                {espula.observacoes && (
                  <p className="text-base text-gray-400 p-3 bg-gray-800/50 rounded">
                    <span className="font-medium">Observações:</span> {espula.observacoes}
                  </p>
                )}
              </CardContent>
            </Card>
          ))}
          
          {finalizedEspulas.length === 0 && (
            <Card className="card-merco-large">
              <CardContent className="p-6 text-center">
                <p className="text-gray-400 text-lg">Nenhuma espula finalizada no histórico.</p>
              </CardContent>
            </Card>
          )}
        </div>
      )}

      {/* Dialog para editar máquinas alocadas */}
      <Dialog open={editingMachines !== null} onOpenChange={() => setEditingMachines(null)}>
        <DialogContent className="dialog-merco max-w-3xl">
          <DialogHeader>
            <DialogTitle className="text-white">
              Editar Máquinas Alocadas - OS {editingMachines?.numero_os}
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-4 max-h-[60vh] overflow-y-auto">
            <div className="flex justify-between items-center mb-3">
              <Label className="text-lg text-white">Máquinas Alocadas</Label>
              <Button size="sm" onClick={addMachineAllocation} variant="outline">
                + Adicionar Máquina
              </Button>
            </div>
            
            {machineAllocations.length === 0 ? (
              <p className="text-gray-400">Carregando alocações...</p>
            ) : (
              machineAllocations.map((allocation, index) => (
                <Card key={`alloc-${index}-${allocation.machine_code}`} className="card-merco p-4">
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                    <div>
                      <Label className="text-white">Máquina *</Label>
                      <Select
                        value={allocation.machine_code || ""}
                        onValueChange={(value) => updateMachineAllocation(index, 'machine', value)}
                      >
                        <SelectTrigger className="bg-gray-800 text-white">
                          <SelectValue placeholder="Selecione a máquina">
                            {allocation.machine_code || "Selecione a máquina"}
                          </SelectValue>
                        </SelectTrigger>
                        <SelectContent>
                          {machines.map(machine => (
                            <SelectItem key={machine.id} value={machine.code}>
                              {machine.code} - {machine.layout_type === "16_fusos" ? "16 Fusos" : "32 Fusos"}
                            </SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </div>
                    <div>
                      <Label className="text-white">Quantidade (metros) *</Label>
                      <Input
                        className="bg-gray-800 text-white"
                        value={allocation.quantidade || ""}
                        onChange={(e) => updateMachineAllocation(index, 'quantidade', e.target.value)}
                        placeholder="Ex: 1.000"
                        type="text"
                      />
                    </div>
                    <div className="flex items-end">
                      {machineAllocations.length > 1 && (
                        <Button
                          size="sm"
                          variant="ghost"
                          className="text-red-400 hover:text-red-300"
                          onClick={() => removeMachineAllocation(index)}
                        >
                          <Trash2 className="h-4 w-4" />
                        </Button>
                      )}
                    </div>
                  </div>
                </Card>
              ))
            )}
            
            <div className="flex gap-2 mt-6">
              <Button onClick={saveMachineAllocations} className="flex-1 btn-merco">
                Salvar Alterações
              </Button>
              <Button onClick={() => setEditingMachines(null)} variant="outline" className="flex-1">
                Cancelar
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>
    </div>
  );
};

const MaintenancePanel = ({ maintenances, user, onMaintenanceUpdate, onMachineUpdate }) => {
  const finishMaintenance = async (maintenanceId) => {
    try {
      await axios.put(`${API}/maintenance/${maintenanceId}/finish`, {}, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Manutenção finalizada com sucesso!");
      onMaintenanceUpdate();
      onMachineUpdate();
    } catch (error) {
      toast.error("Erro ao finalizar manutenção");
    }
  };

  // CORRIGIR FORMATAÇÃO DE HORÁRIO - CONVERSÃO CORRETA UTC PARA BRASÍLIA  
  const formatDateTimeBrazil = (utcString) => {
    if (!utcString) return "-";
    try {
      // Parse UTC string e converte para horário de Brasília
      const utcDate = new Date(utcString + (utcString.includes('Z') ? '' : 'Z'));
      
      return new Intl.DateTimeFormat('pt-BR', {
        timeZone: 'America/Sao_Paulo',
        day: '2-digit',
        month: '2-digit', 
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false
      }).format(utcDate);
    } catch (error) {
      console.error('Error formatting date:', error);
      return "-";
    }
  };

  return (
    <div className="space-y-6">
      <h2 className="text-3xl font-bold text-white">Gerenciamento de Manutenção</h2>
      <div className="grid gap-6">
        {maintenances.map((maintenance) => (
          <Card key={maintenance.id} className="card-merco-large">
            <CardContent className="p-6">
              <div className="flex justify-between items-start mb-4">
                <div>
                  <h3 className="font-bold text-white text-xl">
                    Máquina {maintenance.machine_code}
                  </h3>
                  <p className="text-gray-400 text-lg">Criado por: <span className="text-white">{maintenance.created_by}</span></p>
                </div>
                <Badge className={`${maintenance.status === "em_manutencao" ? "bg-blue-600 text-blue-100" : "bg-green-600 text-green-100"} font-semibold text-sm`}>
                  {maintenance.status === "em_manutencao" ? "EM MANUTENÇÃO" : "FINALIZADA"}
                </Badge>
              </div>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-base mb-4 p-4 bg-gray-900/50 rounded-lg">
                <div className="flex items-center space-x-3">
                  <Clock className="h-5 w-5 text-blue-400" />
                  <div>
                    <span className="font-medium text-gray-400">Iniciado:</span>
                    <p className="text-white">{formatDateTimeBrazil(maintenance.created_at)}</p>
                  </div>
                </div>
                <div className="flex items-center space-x-3">
                  <Clock className="h-5 w-5 text-green-400" />
                  <div>
                    <span className="font-medium text-gray-400">Finalizado:</span>
                    <p className="text-white">{formatDateTimeBrazil(maintenance.finished_at)}</p>
                  </div>
                </div>
              </div>
              
              <p className="text-base text-blue-400 mb-4 p-4 bg-blue-900/20 rounded border border-blue-700">
                <span className="font-medium">Motivo:</span> {maintenance.motivo}
              </p>

              {maintenance.finished_by && (
                <p className="text-base text-gray-400 mb-4">
                  <span className="font-medium">Finalizado por:</span> {maintenance.finished_by}
                </p>
              )}
              
              {maintenance.status === "em_manutencao" && (
                <Button
                  size="lg"
                  className="bg-green-600 hover:bg-green-700 text-white"
                  onClick={() => finishMaintenance(maintenance.id)}
                >
                  Liberar da Manutenção
                </Button>
              )}
            </CardContent>
          </Card>
        ))}
      </div>
    </div>
  );
};

// Banco de Dados Panel
const BancoDadosPanel = ({ user }) => {
  const [artigos, setArtigos] = useState([]);
  const [showForm, setShowForm] = useState(false);
  const [editingArtigo, setEditingArtigo] = useState(null);
  const [formData, setFormData] = useState({
    artigo: '',
    engrenagem: '',
    fios: '',
    maquinas: '',
    ciclos: '',
    carga: ''
  });

  useEffect(() => {
    loadArtigos();
    // Polling para atualização em tempo real a cada 5 segundos
    const interval = setInterval(loadArtigos, 5000);
    return () => clearInterval(interval);
  }, []);

  const loadArtigos = async () => {
    try {
      const response = await axios.get(`${API}/banco-dados`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      setArtigos(response.data);
    } catch (error) {
      console.error("Erro ao carregar artigos:", error);
    }
  };

  const openForm = (artigo = null) => {
    if (artigo) {
      setEditingArtigo(artigo);
      setFormData({
        artigo: artigo.artigo,
        engrenagem: artigo.engrenagem,
        fios: artigo.fios,
        maquinas: artigo.maquinas,
        ciclos: artigo.ciclos || '',
        carga: artigo.carga || ''
      });
    } else {
      setEditingArtigo(null);
      setFormData({
        artigo: '',
        engrenagem: '',
        fios: '',
        maquinas: '',
        ciclos: '',
        carga: ''
      });
    }
    setShowForm(true);
  };

  const saveArtigo = async () => {
    try {
      if (editingArtigo) {
        // Editar
        await axios.put(
          `${API}/banco-dados/${editingArtigo.id}`,
          formData,
          { headers: { Authorization: `Bearer ${localStorage.getItem("token")}` } }
        );
        toast.success("Artigo atualizado com sucesso!");
      } else {
        // Criar
        await axios.post(
          `${API}/banco-dados`,
          formData,
          { headers: { Authorization: `Bearer ${localStorage.getItem("token")}` } }
        );
        toast.success("Artigo criado com sucesso!");
      }
      setShowForm(false);
      loadArtigos();
    } catch (error) {
      console.error("Erro ao salvar artigo:", error);
      toast.error("Erro ao salvar artigo");
    }
  };

  const deleteArtigo = async (id) => {
    if (window.confirm("Tem certeza que deseja excluir este artigo?")) {
      try {
        await axios.delete(`${API}/banco-dados/${id}`, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        });
        toast.success("Artigo excluído com sucesso!");
        loadArtigos();
      } catch (error) {
        console.error("Erro ao excluir artigo:", error);
        toast.error("Erro ao excluir artigo");
      }
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <h2 className="text-3xl font-bold text-white">Banco de Dados - Artigos</h2>
        <Button onClick={() => openForm()} className="btn-merco">
          + Lançar Artigo
        </Button>
      </div>

      <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-3">
        {artigos.map(artigo => (
          <Card key={artigo.id} className="card-merco">
            <CardHeader>
              <CardTitle className="text-white">{artigo.artigo}</CardTitle>
            </CardHeader>
            <CardContent className="space-y-2">
              <div>
                <span className="text-gray-400">Engrenagem:</span>
                <span className="text-white ml-2">{artigo.engrenagem || '-'}</span>
              </div>
              <div>
                <span className="text-gray-400">Fios:</span>
                <span className="text-white ml-2">{artigo.fios || '-'}</span>
              </div>
              <div>
                <span className="text-gray-400">Máquinas:</span>
                <span className="text-white ml-2">{artigo.maquinas || '-'}</span>
              </div>
              <div>
                <span className="text-gray-400">Ciclos:</span>
                <span className="text-white ml-2">{artigo.ciclos || '-'}</span>
              </div>
              <div>
                <span className="text-gray-400">Carga:</span>
                <span className="text-white ml-2">{artigo.carga || '-'}</span>
              </div>
              <div className="flex gap-2 mt-4">
                <Button onClick={() => openForm(artigo)} variant="outline" className="flex-1 bg-blue-600 hover:bg-blue-700 text-white border-blue-600">
                  Editar
                </Button>
                <Button onClick={() => deleteArtigo(artigo.id)} variant="destructive" className="flex-1">
                  Excluir
                </Button>
              </div>
            </CardContent>
          </Card>
        ))}
      </div>

      {/* Dialog para criar/editar */}
      <Dialog open={showForm} onOpenChange={setShowForm}>
        <DialogContent className="dialog-merco">
          <DialogHeader>
            <DialogTitle className="text-white">
              {editingArtigo ? 'Editar Artigo' : 'Novo Artigo'}
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-4">
            <div>
              <Label>Artigo *</Label>
              <Input
                value={formData.artigo}
                onChange={(e) => setFormData({...formData, artigo: e.target.value})}
                placeholder="Nome ou código do artigo"
              />
            </div>
            <div>
              <Label>Engrenagem</Label>
              <Input
                value={formData.engrenagem}
                onChange={(e) => setFormData({...formData, engrenagem: e.target.value})}
                placeholder="Engrenagem"
              />
            </div>
            <div>
              <Label>Fios</Label>
              <Input
                value={formData.fios}
                onChange={(e) => setFormData({...formData, fios: e.target.value})}
                placeholder="Quantidade de fios"
              />
            </div>
            <div>
              <Label>Máquinas</Label>
              <Input
                value={formData.maquinas}
                onChange={(e) => setFormData({...formData, maquinas: e.target.value})}
                placeholder="Máquinas recomendadas"
              />
            </div>
            <div>
              <Label>Ciclos *</Label>
              <Input
                value={formData.ciclos}
                onChange={(e) => setFormData({...formData, ciclos: e.target.value})}
                placeholder="Ciclos"
                required
              />
            </div>
            <div>
              <Label>Carga *</Label>
              <Input
                value={formData.carga}
                onChange={(e) => setFormData({...formData, carga: e.target.value})}
                placeholder="Carga"
                required
              />
            </div>
            <div className="flex gap-2">
              <Button onClick={saveArtigo} className="flex-1 btn-merco" disabled={!formData.artigo || !formData.ciclos || !formData.carga}>
                {editingArtigo ? 'Atualizar' : 'Criar'}
              </Button>
              <Button onClick={() => setShowForm(false)} variant="outline" className="flex-1">
                Cancelar
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>
    </div>
  );
};


const AdminPanel = ({ users, onUserUpdate }) => {
  const [newUser, setNewUser] = useState({
    username: "",
    email: "",
    password: "",
    role: "",
    permissions: {
      dashboard: true,
      producao: true,
      ordem_producao: true,
      relatorios: true,
      espulagem: true,
      manutencao: true,
      banco_dados: true,
      administracao: false
    }
  });

  const [editingUser, setEditingUser] = useState(null);
  const [showEditDialog, setShowEditDialog] = useState(false);
  const [editUser, setEditUser] = useState({
    username: "",
    email: "",
    password: "",
    role: "",
    active: true,
    permissions: {
      dashboard: true,
      producao: true,
      ordem_producao: true,
      relatorios: true,
      espulagem: true,
      manutencao: true,
      banco_dados: true,
      administracao: false
    }
  });

  const togglePermission = (permission) => {
    setNewUser({
      ...newUser,
      permissions: {
        ...newUser.permissions,
        [permission]: !newUser.permissions[permission]
      }
    });
  };

  const toggleEditPermission = (permission) => {
    setEditUser({
      ...editUser,
      permissions: {
        ...editUser.permissions,
        [permission]: !editUser.permissions[permission]
      }
    });
  };

  const openEditDialog = (user) => {
    setEditingUser(user);
    setEditUser({
      username: user.username,
      email: user.email,
      password: "",
      role: user.role,
      active: user.active,
      permissions: user.permissions || {
        dashboard: true,
        producao: true,
        ordem_producao: true,
        relatorios: true,
        espulagem: true,
        manutencao: true,
        banco_dados: true,
        administracao: false
      }
    });
    setShowEditDialog(true);
  };

  const updateUser = async () => {
    try {
      const updateData = {
        username: editUser.username,
        email: editUser.email,
        role: editUser.role,
        active: editUser.active,
        permissions: editUser.permissions
      };
      
      // Only include password if it's not empty
      if (editUser.password && editUser.password.trim() !== "") {
        updateData.password = editUser.password;
      }

      await axios.put(`${API}/users/${editingUser.id}`, updateData, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Usuário atualizado com sucesso!");
      setShowEditDialog(false);
      setEditingUser(null);
      onUserUpdate();
    } catch (error) {
      toast.error(error.response?.data?.detail || "Erro ao atualizar usuário");
    }
  };

  const createUser = async () => {
    try {
      await axios.post(`${API}/users`, newUser, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      toast.success("Usuário criado com sucesso!");
      setNewUser({ 
        username: "", 
        email: "", 
        password: "", 
        role: "",
        permissions: {
          dashboard: true,
          producao: true,
          ordem_producao: true,
          relatorios: true,
          espulagem: true,
          manutencao: true,
          banco_dados: true,
          administracao: false
        }
      });
      onUserUpdate();
    } catch (error) {
      toast.error("Erro ao criar usuário");
    }
  };

  const deleteUser = async (userId) => {
    if (window.confirm("Tem certeza que deseja excluir este usuário?")) {
      try {
        await axios.delete(`${API}/users/${userId}`, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        });
        
        toast.success("Usuário excluído com sucesso!");
        onUserUpdate();
      } catch (error) {
        toast.error("Erro ao excluir usuário");
      }
    }
  };

  const exportReport = async (layoutType) => {
    try {
      const response = await axios.get(`${API}/reports/export?layout_type=${layoutType}`, {
        headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
      });
      
      const data = response.data;
      const wb = XLSX.utils.book_new();
      
      const ordersData = data.orders.map(order => ({
        'ID': order.id,
        'Máquina': order.machine_code,
        'Layout': order.layout_type,
        'Cliente': order.cliente,
        'Artigo': order.artigo,
        'Cor': order.cor,
        'Quantidade': order.quantidade,
        'Status': order.status,
        'Criado por': order.created_by,
        'Criado em': new Date(order.created_at).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'}),
        'Iniciado em': order.started_at ? new Date(order.started_at).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'}) : '',
        'Finalizado em': order.finished_at ? new Date(order.finished_at).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'}) : '',
        'Observação': order.observacao,
        'Obs. Liberação': order.observacao_liberacao,
        'Laudo Final': order.laudo_final
      }));
      
      const ordersWS = XLSX.utils.json_to_sheet(ordersData);
      XLSX.utils.book_append_sheet(wb, ordersWS, "Produção");
      
      const fileName = `relatorio_producao_${layoutType}_${new Date().toISOString().split('T')[0]}.xlsx`;
      XLSX.writeFile(wb, fileName);
      
      toast.success("Relatório exportado com sucesso!");
    } catch (error) {
      toast.error("Erro ao exportar relatório");
    }
  };

  const exportCompleteReport = async () => {
    try {
      // Fetch all data from all endpoints
      const [ordersResponse, ordensResponse, espulasResponse, maintenanceResponse, artigosResponse] = await Promise.all([
        axios.get(`${API}/orders`, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        }),
        axios.get(`${API}/ordens-producao`, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        }),
        axios.get(`${API}/espulas`, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        }),
        axios.get(`${API}/maintenance`, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        }),
        axios.get(`${API}/banco-dados`, {
          headers: { Authorization: `Bearer ${localStorage.getItem("token")}` }
        })
      ]);

      const artigosData = artigosResponse.data || [];
      const wb = XLSX.utils.book_new();
      
      // Sheet 1: Produção (Orders) - com dados do artigo
      const ordersData = ordersResponse.data.map(order => {
        const artigoInfo = artigosData.find(a => 
          a.artigo && order.artigo && 
          a.artigo.toLowerCase() === order.artigo.toLowerCase()
        );
        return {
          'ID': order.id,
          'Número OS': order.numero_os || '-',
          'Origem': order.origem === 'manual' ? 'Manual' : order.origem === 'espulagem' ? 'Espulagem' : 'Ordem',
          'Posição na Fila': order.queue_position || '-',
          'Máquina': order.machine_code,
          'Layout': order.layout_type === '16_fusos' ? '16 Fusos' : '32 Fusos',
          'Cliente': order.cliente,
          'Artigo': order.artigo,
          'Engrenagem': artigoInfo?.engrenagem || '',
          'Fios': artigoInfo?.fios || '',
          'Ciclos': artigoInfo?.ciclos || '',
          'Carga': artigoInfo?.carga || '',
          'Cor': order.cor,
          'Quantidade': order.quantidade,
          'Status': order.status === 'pendente' ? 'Pendente' : order.status === 'em_producao' ? 'Em Produção' : 'Finalizado',
          'Criado por': order.created_by,
          'Criado em': new Date(order.created_at).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'}),
          'Iniciado em': order.started_at ? new Date(order.started_at).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'}) : '',
          'Finalizado em': order.finished_at ? new Date(order.finished_at).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'}) : '',
          'Observação': order.observacao,
          'Obs. Liberação': order.observacao_liberacao,
          'Laudo Final': order.laudo_final
        };
      });
      const ordersWS = XLSX.utils.json_to_sheet(ordersData);
      XLSX.utils.book_append_sheet(wb, ordersWS, "Produção");

      // Sheet 2: Ordens de Produção (todas) - com dados do artigo
      const ordensData = ordensResponse.data.map(ordem => {
        const artigoInfo = artigosData.find(a => 
          a.artigo && ordem.artigo && 
          a.artigo.toLowerCase() === ordem.artigo.toLowerCase()
        );
        return {
          'Número OS': ordem.numero_os,
          'Cliente': ordem.cliente,
          'Artigo': ordem.artigo,
          'Engrenagem': ordem.engrenagem || artigoInfo?.engrenagem || '',
          'Fios': ordem.fios || artigoInfo?.fios || '',
          'Máquinas': ordem.maquinas || artigoInfo?.maquinas || '',
          'Ciclos': artigoInfo?.ciclos || '',
          'Carga/Enchimento': artigoInfo?.carga || '',
          'Cor': ordem.cor,
          'Metragem': ordem.metragem,
          'Data Entrega': new Date(ordem.data_entrega).toLocaleDateString('pt-BR'),
          'Observação': ordem.observacao,
          'Status': ordem.status === 'pendente' ? 'Pendente' : ordem.status === 'em_producao' ? 'Em Produção' : 'Finalizado',
          'Criado por': ordem.criado_por,
          'Criado em': new Date(ordem.criado_em).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'})
        };
      });
      const ordensWS = XLSX.utils.json_to_sheet(ordensData);
      XLSX.utils.book_append_sheet(wb, ordensWS, "Ordens de Produção");

      // Sheet 3: Espulagem - com dados do artigo
      const espulasData = espulasResponse.data.map(espula => {
        const artigoInfo = artigosData.find(a => 
          a.artigo && espula.artigo && 
          a.artigo.toLowerCase() === espula.artigo.toLowerCase()
        );
        
        // Lista de máquinas alocadas
        let maquinasList = '';
        if (espula.machine_allocations && espula.machine_allocations.length > 0) {
          maquinasList = espula.machine_allocations
            .map(alloc => `${alloc.machine_code} (${alloc.quantidade})`)
            .join(', ');
        } else if (espula.maquina) {
          maquinasList = espula.maquina;
        }
        
        const row = {
          'OS': espula.numero_os || espula.id.slice(-8),
          'Cliente': espula.cliente,
          'Artigo': espula.artigo,
          'Engrenagem': artigoInfo?.engrenagem || '',
          'Enchimento': artigoInfo?.carga || '',
          'Ciclos': artigoInfo?.ciclos || '',
          'Máquinas Alocadas': maquinasList,
          'Cor': espula.cor,
          'Mat. Prima': espula.mat_prima,
          'Qtde Fios': espula.qtde_fios || artigoInfo?.fios || '',
          'Qtde Metros': espula.quantidade_metros,
          'Carga': espula.carga
        };
        
        // Adicionar cargas/frações dinamicamente
        if (espula.cargas_fracoes && espula.cargas_fracoes.length > 0) {
          espula.cargas_fracoes.forEach((carga, index) => {
            row[`Fração ${index + 1}`] = carga || '';
          });
        }
        
        row['Data Entrega'] = new Date(espula.data_prevista_entrega).toLocaleDateString('pt-BR');
        row['Observações'] = espula.observacoes;
        row['Status'] = espula.status === 'pendente' ? 'Pendente' : espula.status === 'em_producao_aguardando' ? 'Em Produção (Aguardando)' : espula.status === 'producao' ? 'Produção' : 'Finalizado';
        row['Criado por'] = espula.created_by;
        row['Lançado em'] = new Date(espula.created_at).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'});
        row['Iniciado em'] = espula.iniciado_em ? new Date(espula.iniciado_em).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'}) : '';
        row['Finalizado em'] = espula.finalizado_em ? new Date(espula.finalizado_em).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'}) : '';
        
        return row;
      });
      const espulasWS = XLSX.utils.json_to_sheet(espulasData);
      XLSX.utils.book_append_sheet(wb, espulasWS, "Espulagem");

      // Sheet 4: Manutenção
      const maintenanceData = maintenanceResponse.data.map(maint => ({
        'ID': maint.id,
        'Máquina': maint.machine_code,
        'Motivo': maint.motivo,
        'Status': maint.status === 'em_manutencao' ? 'Em Manutenção' : 'Finalizada',
        'Criado por': maint.created_by,
        'Criado em': new Date(maint.created_at).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'}),
        'Finalizado em': maint.finished_at ? new Date(maint.finished_at).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'}) : '',
        'Finalizado por': maint.finished_by || ''
      }));
      const maintenanceWS = XLSX.utils.json_to_sheet(maintenanceData);
      XLSX.utils.book_append_sheet(wb, maintenanceWS, "Manutenção");

      // Sheet 5: Banco de Dados (Artigos)
      const bancoDadosData = artigosData.map(artigo => ({
        'Artigo': artigo.artigo,
        'Engrenagem': artigo.engrenagem || '',
        'Fios': artigo.fios || '',
        'Máquinas': artigo.maquinas || '',
        'Ciclos': artigo.ciclos || '',
        'Carga': artigo.carga || '',
        'Criado em': new Date(artigo.created_at).toLocaleString('pt-BR', {timeZone: 'America/Sao_Paulo'})
      }));
      const bancoDadosWS = XLSX.utils.json_to_sheet(bancoDadosData);
      XLSX.utils.book_append_sheet(wb, bancoDadosWS, "Banco de Dados");
      
      const fileName = `relatorio_completo_MercoTextil_${new Date().toISOString().split('T')[0]}.xlsx`;
      XLSX.writeFile(wb, fileName);
      
      toast.success("Relatório completo exportado com sucesso!");
    } catch (error) {
      console.error('Error exporting complete report:', error);
      toast.error("Erro ao exportar relatório completo");
    }
  };

  return (
    <div className="space-y-6">
      <h2 className="text-3xl font-bold text-white">Painel de Administração</h2>
      
      <div className="grid md:grid-cols-2 gap-6">
        <Card className="card-merco-large">
          <CardHeader>
            <CardTitle className="flex items-center space-x-2 text-white text-xl">
              <Users className="h-6 w-6 text-red-500" />
              <span>Gerenciar Usuários</span>
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-4 form-merco">
            <div>
              <Label>Usuário</Label>
              <Input
                value={newUser.username}
                onChange={(e) => setNewUser({...newUser, username: e.target.value})}
                placeholder="Nome de usuário"
              />
            </div>
            <div>
              <Label>Email</Label>
              <Input
                value={newUser.email}
                onChange={(e) => setNewUser({...newUser, email: e.target.value})}
                placeholder="Email"
              />
            </div>
            <div>
              <Label>Senha</Label>
              <Input
                type="password"
                value={newUser.password}
                onChange={(e) => setNewUser({...newUser, password: e.target.value})}
                placeholder="Senha"
              />
            </div>
            <div>
              <Label>Função</Label>
              <Select onValueChange={(value) => setNewUser({...newUser, role: value})}>
                <SelectTrigger className="bg-black/60 border-gray-600 text-white">
                  <SelectValue placeholder="Selecione a função" />
                </SelectTrigger>
                <SelectContent className="bg-black border-gray-600">
                  <SelectItem value="admin" className="text-white">Administrador</SelectItem>
                  <SelectItem value="operador_interno" className="text-white">Operador Interno</SelectItem>
                  <SelectItem value="operador_externo" className="text-white">Operador Externo</SelectItem>
                </SelectContent>
              </Select>
            </div>
            <div>
              <Label className="mb-3 block text-lg">Permissões por Aba</Label>
              <div className="grid grid-cols-2 gap-3 p-4 bg-black/40 rounded border border-gray-700">
                <label className="flex items-center space-x-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={newUser.permissions.dashboard}
                    onChange={() => togglePermission('dashboard')}
                    className="w-4 h-4"
                  />
                  <span className="text-white text-sm">Dashboard</span>
                </label>
                <label className="flex items-center space-x-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={newUser.permissions.producao}
                    onChange={() => togglePermission('producao')}
                    className="w-4 h-4"
                  />
                  <span className="text-white text-sm">Produção</span>
                </label>
                <label className="flex items-center space-x-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={newUser.permissions.ordem_producao}
                    onChange={() => togglePermission('ordem_producao')}
                    className="w-4 h-4"
                  />
                  <span className="text-white text-sm">Ordem de Produção</span>
                </label>
                <label className="flex items-center space-x-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={newUser.permissions.relatorios}
                    onChange={() => togglePermission('relatorios')}
                    className="w-4 h-4"
                  />
                  <span className="text-white text-sm">Relatórios</span>
                </label>
                <label className="flex items-center space-x-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={newUser.permissions.espulagem}
                    onChange={() => togglePermission('espulagem')}
                    className="w-4 h-4"
                  />
                  <span className="text-white text-sm">Espulagem</span>
                </label>
                <label className="flex items-center space-x-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={newUser.permissions.manutencao}
                    onChange={() => togglePermission('manutencao')}
                    className="w-4 h-4"
                  />
                  <span className="text-white text-sm">Manutenção</span>
                </label>
                <label className="flex items-center space-x-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={newUser.permissions.banco_dados}
                    onChange={() => togglePermission('banco_dados')}
                    className="w-4 h-4"
                  />
                  <span className="text-white text-sm">Banco de Dados</span>
                </label>
                <label className="flex items-center space-x-2 cursor-pointer col-span-2">
                  <input
                    type="checkbox"
                    checked={newUser.permissions.administracao}
                    onChange={() => togglePermission('administracao')}
                    className="w-4 h-4"
                  />
                  <span className="text-white text-sm font-bold">Administração</span>
                </label>
              </div>
            </div>
            <Button onClick={createUser} className="w-full btn-merco-large">
              Criar Usuário
            </Button>
            
            <div className="mt-6">
              <h4 className="font-medium mb-4 text-white text-lg">Usuários Existentes</h4>
              <div className="space-y-3">
                {users.map((user) => (
                  <div key={user.id} className="flex justify-between items-center p-4 bg-black/40 rounded-lg border border-gray-700">
                    <div className="flex-1">
                      <div className="flex items-center space-x-4 mb-2">
                        <span className="text-white text-base font-medium">{user.username}</span>
                        <Badge className={`${
                          user.role === "admin" ? "badge-admin" : 
                          user.role === "operador_interno" ? "badge-interno" : 
                          "badge-externo"
                        } badge-merco`}>
                          {user.role === "admin" ? "Admin" : 
                           user.role === "operador_interno" ? "Interno" : "Externo"}
                        </Badge>
                        {!user.active && <Badge className="bg-gray-600 text-gray-200">Inativo</Badge>}
                      </div>
                      <div className="text-sm text-gray-400">{user.email}</div>
                    </div>
                    <div className="flex space-x-2">
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => openEditDialog(user)}
                        className="text-blue-400 border-blue-600 hover:bg-blue-600 hover:text-white"
                      >
                        <Edit className="h-4 w-4 mr-1" />
                        Editar
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => deleteUser(user.id)}
                        className="text-red-400 border-red-600 hover:bg-red-600 hover:text-white"
                      >
                        <Trash2 className="h-4 w-4" />
                      </Button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </CardContent>
        </Card>

        {/* Dialog de Edição de Usuário */}
        <Dialog open={showEditDialog} onOpenChange={setShowEditDialog}>
          <DialogContent className="dialog-merco max-w-2xl">
            <DialogHeader className="dialog-header">
              <DialogTitle className="dialog-title">Editar Usuário</DialogTitle>
            </DialogHeader>
            <div className="space-y-4 p-6 form-merco max-h-[70vh] overflow-y-auto">
              <div>
                <Label>Nome de Usuário *</Label>
                <Input
                  value={editUser.username}
                  onChange={(e) => setEditUser({...editUser, username: e.target.value})}
                  placeholder="Nome de usuário"
                />
              </div>
              <div>
                <Label>Email *</Label>
                <Input
                  type="email"
                  value={editUser.email}
                  onChange={(e) => setEditUser({...editUser, email: e.target.value})}
                  placeholder="Email"
                />
              </div>
              <div>
                <Label>Nova Senha (deixe em branco para manter a atual)</Label>
                <Input
                  type="password"
                  value={editUser.password}
                  onChange={(e) => setEditUser({...editUser, password: e.target.value})}
                  placeholder="Digite nova senha ou deixe em branco"
                />
              </div>
              <div>
                <Label>Tipo de Usuário *</Label>
                <Select value={editUser.role} onValueChange={(value) => setEditUser({...editUser, role: value})}>
                  <SelectTrigger>
                    <SelectValue placeholder="Selecione o tipo" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="admin">Admin</SelectItem>
                    <SelectItem value="operador_interno">Operador Interno</SelectItem>
                    <SelectItem value="operador_externo">Operador Externo</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              
              <div>
                <Label className="mb-3 block">Status do Usuário</Label>
                <label className="flex items-center space-x-2 cursor-pointer p-3 bg-black/40 rounded border border-gray-700">
                  <input
                    type="checkbox"
                    checked={editUser.active}
                    onChange={(e) => setEditUser({...editUser, active: e.target.checked})}
                    className="w-4 h-4"
                  />
                  <span className="text-white text-sm font-medium">Usuário Ativo</span>
                </label>
              </div>

              <div>
                <Label className="mb-3 block text-lg">Permissões por Aba</Label>
                <div className="grid grid-cols-2 gap-3 p-4 bg-black/40 rounded border border-gray-700">
                  <label className="flex items-center space-x-2 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={editUser.permissions.dashboard}
                      onChange={() => toggleEditPermission('dashboard')}
                      className="w-4 h-4"
                    />
                    <span className="text-white text-sm">Dashboard</span>
                  </label>
                  <label className="flex items-center space-x-2 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={editUser.permissions.producao}
                      onChange={() => toggleEditPermission('producao')}
                      className="w-4 h-4"
                    />
                    <span className="text-white text-sm">Produção</span>
                  </label>
                  <label className="flex items-center space-x-2 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={editUser.permissions.ordem_producao}
                      onChange={() => toggleEditPermission('ordem_producao')}
                      className="w-4 h-4"
                    />
                    <span className="text-white text-sm">Ordem de Produção</span>
                  </label>
                  <label className="flex items-center space-x-2 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={editUser.permissions.relatorios}
                      onChange={() => toggleEditPermission('relatorios')}
                      className="w-4 h-4"
                    />
                    <span className="text-white text-sm">Relatórios</span>
                  </label>
                  <label className="flex items-center space-x-2 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={editUser.permissions.espulagem}
                      onChange={() => toggleEditPermission('espulagem')}
                      className="w-4 h-4"
                    />
                    <span className="text-white text-sm">Espulagem</span>
                  </label>
                  <label className="flex items-center space-x-2 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={editUser.permissions.manutencao}
                      onChange={() => toggleEditPermission('manutencao')}
                      className="w-4 h-4"
                    />
                    <span className="text-white text-sm">Manutenção</span>
                  </label>
                  <label className="flex items-center space-x-2 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={editUser.permissions.banco_dados}
                      onChange={() => toggleEditPermission('banco_dados')}
                      className="w-4 h-4"
                    />
                    <span className="text-white text-sm">Banco de Dados</span>
                  </label>
                  <label className="flex items-center space-x-2 cursor-pointer col-span-2">
                    <input
                      type="checkbox"
                      checked={editUser.permissions.administracao}
                      onChange={() => toggleEditPermission('administracao')}
                      className="w-4 h-4"
                    />
                    <span className="text-white text-sm font-bold">Administração</span>
                  </label>
                </div>
              </div>

              <div className="flex space-x-3 pt-4">
                <Button 
                  onClick={updateUser} 
                  className="flex-1 btn-merco"
                  disabled={!editUser.username || !editUser.email || !editUser.role}
                >
                  Salvar Alterações
                </Button>
                <Button 
                  variant="outline" 
                  onClick={() => setShowEditDialog(false)} 
                  className="flex-1"
                >
                  Cancelar
                </Button>
              </div>
            </div>
          </DialogContent>
        </Dialog>

        <Card className="card-merco-large">
          <CardHeader>
            <CardTitle className="flex items-center space-x-2 text-white text-xl">
              <Download className="h-6 w-6 text-red-500" />
              <span>Relatórios Excel</span>
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-6">
            <div>
              <h3 className="text-white font-medium mb-3 text-lg">Relatório Completo</h3>
              <Button
                className="w-full bg-blue-600 hover:bg-blue-700 text-white h-auto py-3 px-4 text-base font-semibold"
                onClick={exportCompleteReport}
              >
                <Download className="h-5 w-5 mr-2 flex-shrink-0" />
                <span className="text-left leading-tight">Exportar Relatório Completo<br/><span className="text-xs font-normal opacity-90">Produção, Ordem, Espulagem, Manutenção</span></span>
              </Button>
            </div>
            
            <div className="border-t border-gray-700 pt-4">
              <h3 className="text-white font-medium mb-3 text-lg">Relatórios Individuais de Produção</h3>
              <div className="space-y-3">
                <Button
                  className="w-full bg-green-600 hover:bg-green-700 text-white h-12"
                  onClick={() => exportReport("16_fusos")}
                >
                  <Download className="h-5 w-5 mr-2" />
                  Relatório Produção 16 Fusos
                </Button>
                <Button
                  className="w-full bg-green-600 hover:bg-green-700 text-white h-12"
                  onClick={() => exportReport("32_fusos")}
                >
                  <Download className="h-5 w-5 mr-2" />
                  Relatório Produção 32 Fusos
                </Button>
              </div>
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
};

export default App;