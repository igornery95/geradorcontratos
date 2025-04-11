from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from docx import Document
import os


def gerar_contrato(nome, empresa, cpf, endereco, pasta_selecionada):
    doc = Document()
    texto_contrato = f"""
    CONTRATO DE PRESTAÇÃO DE SERVIÇOS

    Pelo presente instrumento, {nome}, {empresa} portador do CPF {cpf}, residente no endereço {endereco},
    firma o seguinte contrato nos termos descritos a seguir...
    """
    doc.add_paragraph(texto_contrato.strip())
    nome_arquivo = f"contrato_{nome.replace(' ', '_')}.docx"
    caminho_completo = os.path.join(pasta_selecionada, nome_arquivo)
    doc.save(caminho_completo)
    print(f"Contrato salvo em: {caminho_completo}")


class RootWidget(BoxLayout):
    def selecionar_pasta(self):
        content = FileChooserListView(path='/', dirselect=True)
        popup = Popup(title="Selecione a pasta", content=content, size_hint=(0.9, 0.9))

        def selecionar(instance, *args):
            if content.selection:
                self.ids.pasta_label.text = content.selection[0]
                popup.dismiss()

        selecionar_btn = Button(text="Selecionar", size_hint=(1, 0.1))
        selecionar_btn.bind(on_release=selecionar)

        layout = BoxLayout(orientation='vertical')
        layout.add_widget(content)
        layout.add_widget(selecionar_btn)

        popup.content = layout
        popup.open()

    def gerar(self):
        nome = self.ids.nome_input.text
        empresa = self.ids.empresa_input.text
        cpf = self.ids.cpf_input.text
        endereco = self.ids.endereco_input.text
        pasta = self.ids.pasta_label.text

        if nome and empresa and cpf and endereco and pasta:
            gerar_contrato(nome, empresa, cpf, endereco, pasta)
            App.get_running_app().stop()
        else:
            print("Preencha todos os campos e selecione a pasta.")


class GeradorContratoApp(App):
    def build(self):
        return RootWidget()


if __name__ == "__main__":
    GeradorContratoApp().run()
