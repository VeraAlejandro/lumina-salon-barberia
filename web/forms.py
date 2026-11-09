from datetime import date, timedelta

from django import forms

from .models import Reserva


class ReservaForm(forms.ModelForm):
    """Formulario público de reserva con validaciones de negocio."""

    class Meta:
        model = Reserva
        fields = (
            "nombre", "telefono", "correo", "sucursal",
            "servicio", "estilista", "fecha", "hora", "mensaje",
        )
        widgets = {
            "nombre": forms.TextInput(attrs={"placeholder": "Nombre completo"}),
            "telefono": forms.TextInput(attrs={"placeholder": "10 dígitos"}),
            "correo": forms.EmailInput(attrs={"placeholder": "correo@ejemplo.com"}),
            "fecha": forms.DateInput(attrs={"type": "date"}),
            "hora": forms.TimeInput(attrs={"type": "time"}),
            "mensaje": forms.Textarea(
                attrs={"rows": 3, "placeholder": "Cuéntanos si tienes alguna preferencia"}
            ),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        manana = date.today() + timedelta(days=1)
        self.fields["fecha"].widget.attrs.update({"min": manana.isoformat()})
        self.fields["estilista"].required = False
        self.fields["estilista"].empty_label = "Cualquier estilista disponible"
        self.fields["sucursal"].empty_label = "Selecciona una sucursal"
        self.fields["servicio"].empty_label = "Selecciona un servicio"
        for campo in self.fields.values():
            clase = campo.widget.attrs.get("class", "")
            campo.widget.attrs["class"] = f"{clase} campo-formulario".strip()

    def clean_fecha(self):
        fecha = self.cleaned_data["fecha"]
        if fecha <= date.today():
            raise forms.ValidationError("La cita debe agendarse a partir de mañana.")
        return fecha

    def clean(self):
        datos = super().clean()
        sucursal = datos.get("sucursal")
        estilista = datos.get("estilista")
        if estilista and sucursal and estilista.sucursal_id != sucursal.id:
            self.add_error(
                "estilista",
                "El estilista seleccionado pertenece a otra sucursal.",
            )
        return datos
