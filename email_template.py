# Email subject and body template for certificate emails

# Email subject
SUBJECT = "Certificado de Participación - Primer Congreso Latinoamericano de Neurociencias Cognitivas"

# Email body in HTML format
# Use {presenter_name}, {title}, and {authors} as placeholders
BODY = """
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6;">
<p>Buenas <b>{presenter_name}</b>,</p>

<p>Desde el <b>NeuroTransmitiendo</b> queremos expresar nuestro más profundo agradecimiento por haber sido parte de esta primera edición del Primer Congreso Latinoamericano de Neurociencias Cognitivas. La participación activa, el tiempo dedicado y la calidad del trabajo que presentaste hizo de este encuentro un verdadero espacio de intercambio y crecimiento colectivo.</p>

<p>Uno de los principales objetivos del congreso fue crear un lugar de encuentro entre estudiantes, jóvenes investigadores y especialistas consolidados de la región, lo cual fue posible gracias a todas las personas que se sumaron y gracias a vos que compartiste con esta comunidad tu proyecto.</p>

<p>🔗 En el presente correo <b>te adjuntamos el certificado de participación</b> <b>en calidad de presentador</b> del congreso. Por tu presentación: <b>{title}</b> - {authors}</p>

<p>Sabemos que seguir formándose y mantenerse al tanto de los avances en el campo no siempre es fácil. Por eso, desde <b><i>Neurotransmitiendo</i></b> impulsamos propuestas pensadas para quienes buscan seguir profundizándose en estos temas. Una de ellas es nuestra <b>Diplomatura en Neurociencias Cognitivas</b>, un espacio de formación anual online, con clases prácticas y teóricas, dictada por un equipo docente interdisciplinario de primer nivel. Si disfrutaste del congreso, seguro te interese explorar esta propuesta. ¡Está pensada para vos!</p>

<p style="color: red;"><b>⚠️ Como fuiste presentador en el congreso decidimos darte un descuento del 40% OFF ⚠️</b></p>
<p>Con tu descuento personalizado te quedaría en un super precio de 🚨 <b>$</b></p>

<p>👉 Conocé más sobre la diplomatura: <a href="https://www.neurotransmitiendo.org/diplomatura-neurociencias-cognitivas">link</a><br>
👉 Programa completo: <a href="https://docs.google.com/document/d/1u0g8tlnnUToF3Xlp2hLAhSqzyh5KkB-SIvAXtPy6tVk/edit?tab=t.0#heading=h.gjdgxs">link</a></p>

<p>Gracias nuevamente por ser parte de esta experiencia. Nos queda la certeza de que este fue solo el primer paso de muchos por venir para fortalecer la comunidad latinoamericana en neurociencias cognitivas.</p>

<p>Saludos,</p>

<p><i>El equipo de Neurotransmitiendo</i></p>
</body>
</html>
""" 