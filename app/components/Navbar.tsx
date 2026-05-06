"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";

const navItems = [
  { href: "/breakdown", label: "Spinning" },
  { href: "/viscose", label: "Viscose" },
  { href: "/auxilary-recovery", label: "Auxilary&Recovery" },
  { href: "/acid-plant", label: "Acid Plant" },
];

export default function Navbar() {
  const pathname = usePathname();

  return (
    <nav
      style={{
        background: "linear-gradient(90deg, #c0392b 60%, #e67e22 100%)",
        padding: "12px 24px",
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        gap: 16,
      }}
    >
      <Link
        href="/breakdown"
        style={{
          color: "white",
          fontWeight: 700,
          fontSize: 16,
          textDecoration: "none",
          letterSpacing: 0,
        }}
      >
        Birla
      </Link>

      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        {navItems.map((item) => {
          const active =
            pathname === item.href || (pathname === "/" && item.href === "/breakdown");

          return (
            <Link
              key={item.href}
              href={item.href}
              style={{
                color: active ? "#c0392b" : "white",
                background: active ? "white" : "rgba(255,255,255,0.16)",
                border: "1px solid rgba(255,255,255,0.45)",
                borderRadius: 4,
                padding: "7px 12px",
                fontSize: 13,
                fontWeight: 700,
                textDecoration: "none",
              }}
            >
              {item.label}
            </Link>
          );
        })}
      </div>
    </nav>
  );
}
