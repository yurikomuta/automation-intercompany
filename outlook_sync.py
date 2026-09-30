"""Funcoes para sincronizacao do Outlook via COM."""

import time

import pythoncom
import win32com.client


class _EventosSincronizacao:
	concluido = False

	def OnSyncEnd(self):
		self.concluido = True


def sincronizar_todas_as_pastas(namespace, timeout_segundos: int = 300) -> None:
	"""Executa Send/Receive All e aguarda o fim de todos os grupos."""
	sync_objects = namespace.SyncObjects
	eventos = [
		win32com.client.WithEvents(sync_objects.Item(indice), _EventosSincronizacao)
		for indice in range(1, sync_objects.Count + 1)
	]
	if not eventos:
		raise RuntimeError("Nenhum grupo de sincronizacao do Outlook foi encontrado.")

	namespace.SendAndReceive(False)
	limite = time.monotonic() + timeout_segundos
	while not all(evento.concluido for evento in eventos):
		pythoncom.PumpWaitingMessages()
		if time.monotonic() >= limite:
			raise TimeoutError("A sincronizacao do Outlook excedeu o tempo limite.")
		time.sleep(0.1)