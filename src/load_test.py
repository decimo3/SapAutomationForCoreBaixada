''' Module to test SAP_BOT with multiples requests '''
import asyncio
from asyncio import Semaphore, Lock
from solicitacoes import Solicitacao

MAX_INSTANCIAS = 4 # Limite de processos simultâneos

def recuperar_solicitacoes():
	''' Simula acesso ao banco de dados '''
	solicitacoes: list[Solicitacao] = [
		Solicitacao("fatura", 413468215),
		Solicitacao("fatura", 430000711),
		Solicitacao("fatura", 430048243),
		Solicitacao("fatura", 414833022),
		Solicitacao("fatura", 412734253),
		Solicitacao("fatura", 412749369),
		Solicitacao("fatura", 413546217),
		Solicitacao("fatura", 412028039),
		Solicitacao("fatura", 421250890),
		Solicitacao("fatura", 411623182),
		Solicitacao("fatura", 413782839),
		Solicitacao("fatura", 411100113),
		Solicitacao("fatura", 413830209),
		Solicitacao("fatura", 420317511),
		Solicitacao("fatura", 413844717),
		Solicitacao("fatura", 414602608),
		Solicitacao("fatura", 430292103),
		Solicitacao("fatura", 414502014),
	]
	return solicitacoes

async def cooker(solicitacao: Solicitacao):
	''' Simula a função Cooker que roda algo no shell '''
	comando = f"python main.py {solicitacao.aplicacao} {solicitacao.informacao} {solicitacao.instancia}"
	processo = await asyncio.create_subprocess_shell(
		comando,
		stdout=asyncio.subprocess.PIPE,
		stderr=asyncio.subprocess.PIPE
	)
	stdout, stderr = await processo.communicate()
	if processo.returncode != 0:
		print(f"[ERRO] - {stderr.decode().strip()}")
	else:
		print(f"[DONE] - {solicitacao.aplicacao} {solicitacao.informacao} {solicitacao.instancia} - {stdout.decode().strip()}")

# Função principal
async def main():
	# await asyncio.sleep(1)  # Delay inicial
	solicitacoes = recuperar_solicitacoes()

	if not solicitacoes:
		return

	tasks = []
	semaphore = asyncio.Semaphore(MAX_INSTANCIAS)


	instance_control = [False] * MAX_INSTANCIAS
	instance_lock = asyncio.Lock()

	for solicitacao in solicitacoes:
		await semaphore.acquire()

		async with instance_lock:
			instance_number = -1
			for i in range(len(instance_control)):
				if not instance_control[i]:
					instance_control[i] = True
					instance_number = i + 1
					break

		if instance_number == -1:
			semaphore.release()
			continue

		solicitacao.instancia = instance_number

		tasks.append(
			asyncio.create_task(
				executar_com_controle(
					solicitacao, semaphore, instance_control, instance_lock)))

	await asyncio.gather(*tasks)

# Função para controlar execução e liberação de recursos
async def executar_com_controle(
		solicitacao: Solicitacao,
		semaphore: Semaphore,
		instance_control: list[bool],
		instance_lock: Lock):
	print(f"Solicitacao: {solicitacao.aplicacao} {solicitacao.informacao} {solicitacao.instancia}")
	try:
		await cooker(solicitacao)
	finally:
		async with instance_lock:
			instance_index = solicitacao.instancia - 1
			instance_control[instance_index] = False
		semaphore.release()

asyncio.run(main())
