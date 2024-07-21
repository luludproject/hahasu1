var jsonldData = {
  "@context": "https://schema.org",
  "@type": "Organization",
  url: "https://jianhomecare.com",
  logo: "https://jianhomecare.com/img/변기막힘하수구막힘.png",
};

var script = document.createElement("script");
script.type = "application/ld+json";
script.text = JSON.stringify(jsonldData);
document.head.appendChild(script);
