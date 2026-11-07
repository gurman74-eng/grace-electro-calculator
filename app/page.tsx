"use client";

import { useMemo, useState } from "react";

type Tier = "Розница" | "Опт" | "Крупный опт";

const items = [
  { name: "Плинтус напольный алюминиевый, малый, 20×80×2000 мм", unit: "шт.", pack: 2, prices: [680, 525, 370] },
  { name: "Плинтус напольный алюминиевый, большой, 20×80×2000 мм", unit: "шт.", pack: 2, prices: [980, 765, 545] },
  { name: "Плинтус напольный алюминиевый, без покрытия, 20×80×2000 мм", unit: "шт.", pack: 2, prices: [790, 600, 425] },
  { name: "Соединительный элемент для плинтуса, малый", unit: "шт.", prices: [2320, 1800, 1275] },
  { name: "Соединительный элемент для плинтуса, большой", unit: "шт.", prices: [2450, 1900, 1350] },
  { name: "Алюминиевый угол 2 м, 45×45", unit: "шт.", prices: [2100, 1650, 1160] },
  { name: "Внешний угол декоративный, малый", unit: "шт.", prices: [575, 450, 320] },
  { name: "Внешний угол декоративный, большой", unit: "шт.", prices: [625, 500, 350] },
  { name: "Заглушка, малая", unit: "шт.", prices: [360, 275, 190] },
  { name: "Заглушка, большая", unit: "шт.", prices: [410, 300, 210] },
  { name: "Угол внутренний 20×80 мм, малый", unit: "шт.", prices: [735, 570, 400] },
  { name: "Угол внутренний 20×80 мм, большой", unit: "шт.", prices: [785, 595, 420] },
];

const money = new Intl.NumberFormat("ru-RU", { style: "currency", currency: "RUB", maximumFractionDigits: 0 });
const esc = (value: string | number) => String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

export default function Home() {
  const [meters, setMeters] = useState(100);
  const [quantities, setQuantities] = useState<number[]>([50, 0, ...Array(items.length - 2).fill(0)]);
  const tierIndex = meters <= 500 ? 0 : meters <= 1000 ? 1 : 2;
  const tier: Tier = ["Розница", "Опт", "Крупный опт"][tierIndex] as Tier;
  const rows = useMemo(() => items.map((item, index) => ({ ...item, quantity: quantities[index] || 0, price: item.prices[tierIndex], total: (quantities[index] || 0) * item.prices[tierIndex] })), [quantities, tierIndex]);
  const total = rows.reduce((sum, row) => sum + row.total, 0);
  const setQuantity = (index: number, value: string) => { const next = [...quantities]; next[index] = Math.max(0, Number(value) || 0); setQuantities(next); };
  const downloadExcel = () => {
    const active = rows.filter((row) => row.quantity > 0);
    const xml = `<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="Расчёт"><Table><Row>${["Позиция", "Количество", "Ед.", "Тариф", "Цена, ₽", "Сумма, ₽"].map(h => `<Cell><Data ss:Type="String">${h}</Data></Cell>`).join("")}</Row>${active.map(r => `<Row><Cell><Data ss:Type="String">${esc(r.name)}</Data></Cell><Cell><Data ss:Type="Number">${r.quantity}</Data></Cell><Cell><Data ss:Type="String">${r.unit}</Data></Cell><Cell><Data ss:Type="String">${tier}</Data></Cell><Cell><Data ss:Type="Number">${r.price}</Data></Cell><Cell><Data ss:Type="Number">${r.total}</Data></Cell></Row>`).join("")}<Row><Cell ss:MergeAcross="4"><Data ss:Type="String">Итого</Data></Cell><Cell><Data ss:Type="Number">${total}</Data></Cell></Row></Table></Worksheet></Workbook>`;
    const url = URL.createObjectURL(new Blob([xml], { type: "application/vnd.ms-excel" })); const a = document.createElement("a"); a.href = url; a.download = `Расчёт_Grace_Electro_${tier}.xls`; a.click(); URL.revokeObjectURL(url);
  };
  return <main>
    <section className="hero"><div className="brand">GRACE <span>ELECTRO</span></div><p className="eyebrow">Плинтус и комплектующие · Санкт-Петербург</p><h1>Стоимость заказа<br/><em>за пару минут.</em></h1><p className="lead">Введите длину основания плинтуса — калькулятор сам применит нужный прайс и подготовит расчёт для Excel.</p></section>
    <section className="calculator" aria-label="Калькулятор стоимости">
      <div className="calc-top"><div><p className="eyebrow">01 · Основание расчёта</p><label htmlFor="meters">Погонаж основания плинтуса</label><div className="meter-input"><input id="meters" type="number" min="0" value={meters} onChange={e => setMeters(Math.max(0, Number(e.target.value) || 0))}/><span>м</span></div></div><div className={`tier tier-${tierIndex}`}><span>Ваш тариф</span><strong>{tier}</strong><small>{tierIndex === 0 ? "до 500 м" : tierIndex === 1 ? "от 501 до 1 000 м" : "свыше 1 000 м"}</small></div></div>
      <div className="rules"><span>до 500 м — <b>розница</b></span><span>501–1 000 м — <b>опт</b></span><span>свыше 1 000 м — <b>крупный опт</b></span></div>
      <div className="products"><div className="products-head"><p className="eyebrow">02 · Состав заказа</p><span>Цены автоматически взяты из прайса «{tier}»</span></div>{rows.map((row, i) => <div className="product" key={row.name}><div><h3>{row.name}</h3>{row.pack && <p>Длина одной планки — {row.pack} м</p>}</div><div className="price">{money.format(row.price)}<small>за {row.unit}</small></div><label className="quantity"><span className="sr-only">Количество</span><input aria-label={`Количество: ${row.name}`} type="number" min="0" step="1" value={row.quantity || ""} placeholder="0" onChange={e => setQuantity(i, e.target.value)}/><small>{row.unit}</small></label><div className="line-total">{money.format(row.total)}</div></div>)}</div>
      <div className="summary"><div><span>Итого к оплате</span><strong>{money.format(total)}</strong><small>НДС и доставка рассчитываются отдельно</small></div><button onClick={downloadExcel}>Скачать расчёт в Excel <span>↓</span></button></div>
    </section><footer><span>Grace Electro · 2026</span><a href="tel:+79111565577">+7 911 156-55-77</a><a href="mailto:grace-electro@mail.ru">grace-electro@mail.ru</a></footer>
  </main>;
}
