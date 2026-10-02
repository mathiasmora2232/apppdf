import { FileSignature, Layers3, ShieldCheck, Users } from "lucide-react";

const cards = [
  { icon: Layers3, title: "Lotes", text: "Procesa cientos de documentos en una sola operación." },
  { icon: Users, title: "Multifirmante", text: "Secuencial, paralelo o flujos mixtos." },
  { icon: FileSignature, title: "Multipágina", text: "Campos normalizados por página y firmante." },
  { icon: ShieldCheck, title: "Auditable", text: "Hashes, eventos y evidencia de cada operación." },
];

export default function Home() {
  return (
    <main>
      <section className="hero">
        <span className="badge">SIGNFLOW · DOCUMENT SIGNING</span>
        <h1>Firma documentos sin convertir tu operación en un infierno de PDFs.</h1>
        <p>
          Plataforma base para firma por lotes, múltiples firmantes y múltiples páginas,
          con motor criptográfico separado.
        </p>
        <div className="actions">
          <button>Nuevo lote</button>
          <button className="secondary">Ver documentos</button>
        </div>
      </section>

      <section className="grid">
        {cards.map(({ icon: Icon, title, text }) => (
          <article key={title}>
            <Icon size={28} />
            <h2>{title}</h2>
            <p>{text}</p>
          </article>
        ))}
      </section>

      <section className="workspace">
        <div>
          <small>LOTE ACTIVO</small>
          <h2>Contratos octubre 2026</h2>
          <p>128 documentos · 3 firmantes · modo secuencial</p>
        </div>
        <strong>72%</strong>
      </section>
    </main>
  );
}
