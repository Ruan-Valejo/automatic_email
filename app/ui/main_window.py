from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QLabel,
    QComboBox,
    QLineEdit,
    QTextEdit,
    QPushButton,
    QFileDialog,
    QListWidget,
    QVBoxLayout,
    QHBoxLayout,
    QMessageBox,
    QGroupBox,
)

from app.db import SessionLocal
from app.models import Client
from app.services.outlook_service import create_draft_email
from app.services.subject_service import generate_subject


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Automatizado")
        self.resize(950, 720)

        self.attachments = []
        self.current_client_id = None

        self._build_ui()
        self.load_clients()

    def _build_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)

        # =========================
        # BLOCO DE CADASTRO CLIENTE
        # =========================
        client_group = QGroupBox("Cadastro de cliente")
        client_layout = QVBoxLayout()
        client_group.setLayout(client_layout)

        select_layout = QHBoxLayout()
        select_label = QLabel("Clientes salvos:")
        self.client_combo = QComboBox()
        self.client_combo.currentIndexChanged.connect(self.on_client_changed)

        self.new_client_button = QPushButton("Novo cliente")
        self.new_client_button.clicked.connect(self.new_client)

        select_layout.addWidget(select_label)
        select_layout.addWidget(self.client_combo)
        select_layout.addWidget(self.new_client_button)

        name_label = QLabel("Nome do cliente:")
        self.client_name_input = QLineEdit()

        to_label = QLabel("Para:")
        self.to_input = QLineEdit()

        cc_label = QLabel("CC:")
        self.cc_input = QTextEdit()
        self.cc_input.setFixedHeight(70)

        self.save_client_button = QPushButton("Salvar cliente")
        self.save_client_button.clicked.connect(self.save_client)

        client_layout.addLayout(select_layout)
        client_layout.addWidget(name_label)
        client_layout.addWidget(self.client_name_input)
        client_layout.addWidget(to_label)
        client_layout.addWidget(self.to_input)
        client_layout.addWidget(cc_label)
        client_layout.addWidget(self.cc_input)
        client_layout.addWidget(self.save_client_button)

        # =========================
        # BLOCO DO E-MAIL
        # =========================
        email_group = QGroupBox("Composição do e-mail")
        email_layout = QVBoxLayout()
        email_group.setLayout(email_layout)

        subject_label = QLabel("Assunto:")
        self.subject_input = QLineEdit()

        subject_layout = QHBoxLayout()
        subject_layout.addWidget(self.subject_input)

        self.generate_subject_button = QPushButton("Gerar assunto")
        self.generate_subject_button.clicked.connect(self.generate_email_subject)
        subject_layout.addWidget(self.generate_subject_button)

        body_label = QLabel("Corpo do e-mail:")
        self.body_input = QTextEdit()

        attachments_label = QLabel("Anexos:")
        self.attachments_list = QListWidget()

        attachments_buttons_layout = QHBoxLayout()
        self.add_attachment_button = QPushButton("Adicionar anexo")
        self.remove_attachment_button = QPushButton("Remover anexo selecionado")

        self.add_attachment_button.clicked.connect(self.add_attachment)
        self.remove_attachment_button.clicked.connect(self.remove_selected_attachment)

        attachments_buttons_layout.addWidget(self.add_attachment_button)
        attachments_buttons_layout.addWidget(self.remove_attachment_button)

        self.create_draft_button = QPushButton("Abrir novo e-mail no Outlook")
        self.create_draft_button.clicked.connect(self.create_outlook_draft)

        email_layout.addWidget(subject_label)
        email_layout.addLayout(subject_layout)
        email_layout.addWidget(body_label)
        email_layout.addWidget(self.body_input)
        email_layout.addWidget(attachments_label)
        email_layout.addWidget(self.attachments_list)
        email_layout.addLayout(attachments_buttons_layout)
        email_layout.addWidget(self.create_draft_button)

        main_layout.addWidget(client_group)
        main_layout.addWidget(email_group)

    def load_clients(self, selected_id=None):
        session = SessionLocal()
        try:
            clients = session.query(Client).order_by(Client.name).all()

            self.client_combo.blockSignals(True)
            self.client_combo.clear()
            self.client_combo.addItem("Selecione um cliente", None)

            selected_index = 0

            for index, client in enumerate(clients, start=1):
                self.client_combo.addItem(client.name, client.id)
                if selected_id is not None and client.id == selected_id:
                    selected_index = index

            self.client_combo.setCurrentIndex(selected_index)
            self.client_combo.blockSignals(False)

            if selected_index == 0:
                self.clear_client_fields()
            else:
                self.on_client_changed(selected_index)

        finally:
            session.close()

    def on_client_changed(self, index):
        client_id = self.client_combo.itemData(index)

        self.clear_email_fields()

        if not client_id:
            self.current_client_id = None
            self.clear_client_fields()
            return

        session = SessionLocal()
        try:
            client = session.get(Client, client_id)
            if client:
                self.current_client_id = client.id
                self.fill_client_data(client)
        finally:
            session.close()

    def fill_client_data(self, client: Client):
            self.client_name_input.setText(client.name or "")
            self.to_input.setText(client.to_email or "")
            self.cc_input.setPlainText(client.cc_emails or "")

    def clear_client_fields(self):
        self.client_name_input.clear()
        self.to_input.clear()
        self.cc_input.clear()

    def new_client(self):
        self.current_client_id = None
        self.client_combo.setCurrentIndex(0)
        self.clear_client_fields()
        self.clear_email_fields()
        self.client_name_input.setFocus()

    def save_client(self):
        name = self.client_name_input.text().strip()
        to_email = self.to_input.text().strip()
        cc_emails = self.cc_input.toPlainText().strip()

        if not name:
            QMessageBox.warning(self, "Aviso", "Informe o nome do cliente.")
            return

        if not to_email:
            QMessageBox.warning(self, "Aviso", "Informe o e-mail do campo 'Para'.")
            return

        session = SessionLocal()
        try:
            # edição de cliente existente
            if self.current_client_id is not None:
                client = session.get(Client, self.current_client_id)

                if client is None:
                    QMessageBox.warning(self, "Aviso", "Cliente não encontrado.")
                    return

                duplicate = (
                    session.query(Client)
                    .filter(Client.name == name, Client.id != self.current_client_id)
                    .first()
                )
                if duplicate:
                    QMessageBox.warning(self, "Aviso", "Já existe outro cliente com esse nome.")
                    return

                client.name = name
                client.to_email = to_email
                client.cc_emails = cc_emails
                session.commit()

                QMessageBox.information(self, "Sucesso", "Cliente salvo com sucesso.")
                self.load_clients(selected_id=client.id)
                return

            # criação de novo cliente
            existing = session.query(Client).filter_by(name=name).first()
            if existing:
                QMessageBox.warning(self, "Aviso", "Já existe um cliente com esse nome.")
                return

            new_client = Client(
                name=name,
                to_email=to_email,
                cc_emails=cc_emails,
            )
            session.add(new_client)
            session.commit()
            session.refresh(new_client)

            self.current_client_id = new_client.id

            QMessageBox.information(self, "Sucesso", "Cliente salvo com sucesso.")
            self.load_clients(selected_id=new_client.id)

        finally:
            session.close()

    def add_attachment(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecionar anexos"
        )

        if not files:
            return

        for file_path in files:
            if file_path not in self.attachments:
                self.attachments.append(file_path)
                self.attachments_list.addItem(file_path)

        if self.client_name_input.text().strip():
             self.subject_input.setText(
                generate_subject(self.client_name_input.text().strip(), self.attachments)
    )

    def remove_selected_attachment(self):
        current_row = self.attachments_list.currentRow()

        if current_row < 0:
            QMessageBox.warning(self, "Aviso", "Selecione um anexo para remover.")
            return

        self.attachments.pop(current_row)
        self.attachments_list.takeItem(current_row)

        if self.client_name_input.text().strip():
            self.subject_input.setText(
              generate_subject(self.client_name_input.text().strip(), self.attachments)
    )

    def generate_email_subject(self):
        client_name = self.client_name_input.text().strip()

        if not client_name:
            QMessageBox.warning(self, "Aviso", "Selecione ou informe um cliente antes de gerar o assunto.")
            return

        subject = generate_subject(client_name, self.attachments)
        self.subject_input.setText(subject)

    def create_outlook_draft(self):
        to_email = self.to_input.text().strip()
        cc_email = self.cc_input.toPlainText().strip()
        subject = self.subject_input.text().strip()
        body = self.body_input.toPlainText().strip()

        if not to_email:
            QMessageBox.warning(self, "Aviso", "O campo 'Para' é obrigatório.")
            return

        try:
            create_draft_email(
                to_email=to_email,
                cc_email=cc_email,
                subject=subject,
                body=body,
                attachments=self.attachments,
            )
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível abrir o e-mail no Outlook.\n\n{e}")

    def clear_email_fields(self):
        self.subject_input.clear()
        self.body_input.clear()
        self.attachments.clear()
        self.attachments_list.clear()