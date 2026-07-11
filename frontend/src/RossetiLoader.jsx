// Заставка загрузки: надпись «РОССЕТИ», буквы загораются по очереди
// (электрик-синий с glow). Перенесена 1:1 из «Учёта ПУ».
import React from "react";

export default function RossetiLoader({ size = "normal" }) {
  const letters = ["Р", "О", "С", "С", "Е", "Т", "И"];
  return (
    <div className={"rosseti-loader" + (size === "small" ? " rosseti-loader--small" : "")}>
      {letters.map((letter, idx) => (
        <span key={idx} className="rosseti-letter" style={{ animationDelay: `${idx * 0.3}s` }}>
          {letter}
        </span>
      ))}
    </div>
  );
}
