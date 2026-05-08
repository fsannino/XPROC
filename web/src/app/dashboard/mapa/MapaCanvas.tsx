'use client'

import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { useRouter } from 'next/navigation'
import {
  ReactFlow,
  ReactFlowProvider,
  Background,
  Controls,
  MiniMap,
  type Node,
  type Edge,
  type NodeChange,
  applyNodeChanges,
  type OnNodesChange,
} from '@xyflow/react'
import '@xyflow/react/dist/style.css'
import type { MapaNode, MapaEdge } from '@/actions/mapa'
import { upsertNode, updateNodePosition, deleteNode } from '@/actions/mapa'
import {
  getTransacoesDeProcesso,
  getTransacoesDeAtividade,
  setProcessoTransacoes,
  setAtividadeTransacoes,
} from '@/actions/relacoes'
import { getRaciDoProcesso, setRaciDoProcesso } from '@/actions/raci'
import { getProdutosDoProcesso, setProdutosDoProcesso } from '@/actions/produtos'
import {
  getInsumosDeProcesso,
  setInsumosDeProcesso,
  getInsumosDeAtividade,
  setInsumosDeAtividade,
} from '@/actions/insumos'
import { getSistemasDoProcesso, setSistemasDoProcesso } from '@/actions/sistemas'
import {
  listarDependenciasDeProcesso,
  listarDependenciasDeAtividade,
  criarDependenciaProcesso,
  criarDependenciaAtividade,
  excluirDependenciaProcesso,
  excluirDependenciaAtividade,
  type DependenciaProcessoView,
  type DependenciaAtividadeView,
} from '@/actions/dependencias'
import { MAPA_LEVELS, type NodeType, type DependenciaTipo } from '@/lib/definitions'
import NodeDrawer, { type DrawerState } from '@/components/mapa/NodeDrawer'
import CustomNode from '@/components/mapa/CustomNode'
import type { MultiSelectOption } from '@/components/ui/multi-select'
import type { RaciAtribuicao, PessoaOption } from '@/components/mapa/RaciSection'
import type { ProdutoOption } from '@/components/mapa/ProdutosSection'
import type { InsumoOption, InsumoVinculo } from '@/components/mapa/InsumosSection'
import type { SistemaOption, SistemaVinculo } from '@/components/mapa/SistemasSection'
import type { DependenciaItem, NoOption } from '@/components/mapa/DependenciasSection'

type Props = {
  initialNodes: MapaNode[]
  initialEdges: MapaEdge[]
  transacoes: MultiSelectOption[]
  pessoas: PessoaOption[]
  produtos: ProdutoOption[]
  insumos: InsumoOption[]
  sistemas: SistemaOption[]
  dependenciasProcesso: DependenciaProcessoView[]
  dependenciasAtividade: DependenciaAtividadeView[]
}

const nodeTypes = { xproc: CustomNode }

function autoLayout(node: MapaNode, index: number): { x: number; y: number } {
  if (node.posicaoX !== 0 || node.posicaoY !== 0) {
    return { x: node.posicaoX, y: node.posicaoY }
  }
  const yByType: Record<NodeType, number> = {
    cadeia: 0, macroprocesso: 180, processo: 360,
    macroatividade: 540, atividade: 720,
  }
  return { x: 60 + index * 260, y: yByType[node.tipo] }
}

function nodeKey(tipo: NodeType, id: number) {
  return `${tipo}:${id}`
}

function parseKey(key: string): { tipo: NodeType; id: number } {
  const [tipo, id] = key.split(':') as [NodeType, string]
  return { tipo, id: Number(id) }
}

export default function MapaCanvas(props: Props) {
  return (
    <ReactFlowProvider>
      <CanvasInner {...props} />
    </ReactFlowProvider>
  )
}

function CanvasInner(props: Props) {
  const {
    initialNodes,
    initialEdges,
    transacoes,
    pessoas,
    produtos,
    insumos,
    sistemas,
    dependenciasProcesso,
    dependenciasAtividade,
  } = props
  const router = useRouter()

  const grouped = useMemo(() => {
    const indexByType: Record<NodeType, number> = {
      cadeia: 0, macroprocesso: 0, processo: 0, macroatividade: 0, atividade: 0,
    }
    return initialNodes.map((n) => {
      const idx = indexByType[n.tipo]++
      const pos = autoLayout(n, idx)
      const node: Node = {
        id: nodeKey(n.tipo, n.id),
        type: 'xproc',
        position: pos,
        data: {
          tipo: n.tipo,
          label: n.descricao,
          abreviacao: n.abreviacao,
          sequencia: n.sequencia,
          tempoMedioCiclo: n.tempoMedioCiclo,
          custoEstimado: n.custoEstimado,
          volumeMensal: n.volumeMensal,
          onAddChild: () => openDrawerFor({ mode: 'create', tipoFilho: childTipo(n.tipo), parentId: n.id }),
          onDelete: () => handleDelete(n.tipo, n.id),
        },
      }
      return node
    })
  }, [initialNodes]) // eslint-disable-line react-hooks/exhaustive-deps

  const initialFlowEdges: Edge[] = useMemo(() => {
    const hierarquicas: Edge[] = initialEdges.map((e) => ({
      id: `e:${e.source}->${e.target}`,
      source: e.source,
      target: e.target,
      type: 'smoothstep',
      style: { stroke: '#1A6E8E', strokeWidth: 1.5 },
    }))
    const depsP: Edge[] = dependenciasProcesso.map((d) => ({
      id: `dp:${d.id}`,
      source: nodeKey('processo', d.origemId),
      target: nodeKey('processo', d.destinoId),
      type: 'smoothstep',
      animated: true,
      style: { stroke: '#F7A823', strokeWidth: 1.5, strokeDasharray: '6 4' },
      label: d.tipo === 'PRECEDE' ? '' : d.tipo.toLowerCase(),
      labelStyle: { fontSize: 10, fill: '#1E2D3D' },
    }))
    const depsA: Edge[] = dependenciasAtividade.map((d) => ({
      id: `da:${d.id}`,
      source: nodeKey('atividade', d.origemId),
      target: nodeKey('atividade', d.destinoId),
      type: 'smoothstep',
      animated: true,
      style: { stroke: '#F7A823', strokeWidth: 1.5, strokeDasharray: '6 4' },
      label: d.tipo === 'PRECEDE' ? '' : d.tipo.toLowerCase(),
      labelStyle: { fontSize: 10, fill: '#1E2D3D' },
    }))
    return [...hierarquicas, ...depsP, ...depsA]
  }, [initialEdges, dependenciasProcesso, dependenciasAtividade])

  const [nodes, setNodes] = useState<Node[]>(grouped)
  const [edges] = useState<Edge[]>(initialFlowEdges)
  const [drawer, setDrawer] = useState<DrawerState | null>(null)
  const [drawerTransacoes, setDrawerTransacoes] = useState<string[] | null>(null)
  const [drawerRaci, setDrawerRaci] = useState<RaciAtribuicao[] | null>(null)
  const [drawerProdutos, setDrawerProdutos] = useState<number[] | null>(null)
  const [drawerInsumos, setDrawerInsumos] = useState<InsumoVinculo[] | null>(null)
  const [drawerSistemas, setDrawerSistemas] = useState<SistemaVinculo[] | null>(null)
  const [drawerDepsSaidas, setDrawerDepsSaidas] = useState<DependenciaItem[] | null>(null)
  const [drawerDepsEntradas, setDrawerDepsEntradas] = useState<DependenciaItem[] | null>(null)

  const dragTimers = useRef(new Map<string, ReturnType<typeof setTimeout>>())

  function resetDrawerData() {
    setDrawerTransacoes(null)
    setDrawerRaci(null)
    setDrawerProdutos(null)
    setDrawerInsumos(null)
    setDrawerSistemas(null)
    setDrawerDepsSaidas(null)
    setDrawerDepsEntradas(null)
  }

  const openDrawerFor = useCallback((s: DrawerState) => {
    setDrawer(s)
    resetDrawerData()
  }, [])

  const handleDelete = useCallback(async (tipo: NodeType, id: number) => {
    if (!confirm(`Excluir ${MAPA_LEVELS[tipo].label.toLowerCase()}?`)) return
    const res = await deleteNode({ tipo, id })
    if (!res.success) {
      alert(res.error ?? 'Falha ao excluir.')
      return
    }
    setNodes((ns) => ns.filter((n) => n.id !== nodeKey(tipo, id)))
    router.refresh()
  }, [router])

  const onNodesChange: OnNodesChange = useCallback((changes: NodeChange[]) => {
    setNodes((ns) => applyNodeChanges(changes, ns))
    for (const ch of changes) {
      if (ch.type !== 'position' || !ch.position || ch.dragging) continue
      const { tipo, id } = parseKey(ch.id)
      const prev = dragTimers.current.get(ch.id)
      if (prev) clearTimeout(prev)
      const t = setTimeout(() => {
        updateNodePosition({
          tipo,
          id,
          posicaoX: ch.position!.x,
          posicaoY: ch.position!.y,
        }).catch(() => undefined)
      }, 250)
      dragTimers.current.set(ch.id, t)
    }
  }, [])

  const onNodeClick = useCallback(async (_: React.MouseEvent, n: Node) => {
    const { tipo, id } = parseKey(n.id)
    const data = n.data as {
      label: string
      abreviacao?: string | null
      sequencia?: number | null
      tempoMedioCiclo?: number | null
      custoEstimado?: number | null
      volumeMensal?: number | null
    }
    setDrawer({
      mode: 'edit',
      tipo,
      id,
      descricao: data.label,
      abreviacao: data.abreviacao ?? '',
      sequencia: data.sequencia ?? undefined,
      tempoMedioCiclo: data.tempoMedioCiclo ?? null,
      custoEstimado: data.custoEstimado ?? null,
      volumeMensal: data.volumeMensal ?? null,
    })
    resetDrawerData()

    if (tipo === 'processo') {
      try {
        const [txIds, raci, produtoIds, insumoVinc, sistemaVinc, deps] = await Promise.all([
          getTransacoesDeProcesso(id),
          getRaciDoProcesso(id),
          getProdutosDoProcesso(id),
          getInsumosDeProcesso(id),
          getSistemasDoProcesso(id),
          listarDependenciasDeProcesso(id),
        ])
        setDrawerTransacoes(txIds)
        setDrawerRaci(raci.map((a) => ({ pessoaId: a.pessoaId, papel: a.papel })))
        setDrawerProdutos(produtoIds.map((p) => p.id))
        setDrawerInsumos(insumoVinc.map((i) => ({ insumoId: i.id, direcao: i.direcao })))
        setDrawerSistemas(sistemaVinc.map((s) => ({ sistemaId: s.id, papel: s.papel })))
        setDrawerDepsSaidas(deps.saidas)
        setDrawerDepsEntradas(deps.entradas)
      } catch {
        setDrawerTransacoes([])
        setDrawerRaci([])
        setDrawerProdutos([])
        setDrawerInsumos([])
        setDrawerSistemas([])
        setDrawerDepsSaidas([])
        setDrawerDepsEntradas([])
      }
    } else if (tipo === 'atividade') {
      try {
        const [txIds, insumoVinc, deps] = await Promise.all([
          getTransacoesDeAtividade(id),
          getInsumosDeAtividade(id),
          listarDependenciasDeAtividade(id),
        ])
        setDrawerTransacoes(txIds)
        setDrawerInsumos(insumoVinc.map((i) => ({ insumoId: i.id, direcao: i.direcao })))
        setDrawerDepsSaidas(deps.saidas)
        setDrawerDepsEntradas(deps.entradas)
      } catch {
        setDrawerTransacoes([])
        setDrawerInsumos([])
        setDrawerDepsSaidas([])
        setDrawerDepsEntradas([])
      }
    }
  }, [])

  const handleSubmit = useCallback(async (payload: {
    mode: 'create' | 'edit'
    tipo: NodeType
    id?: number
    parentId?: number
    descricao: string
    abreviacao?: string
    sequencia?: number
    tempoMedioCiclo?: number
    custoEstimado?: number
    volumeMensal?: number
    transacaoIds?: string[]
    raci?: RaciAtribuicao[]
    produtoIds?: number[]
    insumos?: InsumoVinculo[]
    sistemas?: SistemaVinculo[]
  }) => {
    const res = await upsertNode({
      tipo: payload.tipo,
      id: payload.id,
      parentId: payload.parentId,
      descricao: payload.descricao,
      abreviacao: payload.abreviacao,
      sequencia: payload.sequencia,
      tempoMedioCiclo: payload.tempoMedioCiclo,
      custoEstimado: payload.custoEstimado,
      volumeMensal: payload.volumeMensal,
    })
    if (!res?.success) {
      alert(res?.error ?? 'Falha ao salvar.')
      return false
    }

    const id = payload.id
    if (id && payload.mode === 'edit') {
      const ops: Promise<unknown>[] = []
      if (payload.transacaoIds) {
        if (payload.tipo === 'processo') {
          ops.push(setProcessoTransacoes({ processoId: id, transacaoIds: payload.transacaoIds }))
        } else if (payload.tipo === 'atividade') {
          ops.push(setAtividadeTransacoes({ atividadeId: id, transacaoIds: payload.transacaoIds }))
        }
      }
      if (payload.raci && payload.tipo === 'processo') {
        ops.push(setRaciDoProcesso({ processoId: id, atribuicoes: payload.raci }))
      }
      if (payload.produtoIds && payload.tipo === 'processo') {
        ops.push(setProdutosDoProcesso({ processoId: id, produtoIds: payload.produtoIds }))
      }
      if (payload.insumos) {
        if (payload.tipo === 'processo') {
          ops.push(setInsumosDeProcesso({ processoId: id, vinculos: payload.insumos }))
        } else if (payload.tipo === 'atividade') {
          ops.push(setInsumosDeAtividade({ atividadeId: id, vinculos: payload.insumos }))
        }
      }
      if (payload.sistemas && payload.tipo === 'processo') {
        ops.push(setSistemasDoProcesso({ processoId: id, vinculos: payload.sistemas }))
      }
      if (ops.length > 0) await Promise.all(ops)
    }

    setDrawer(null)
    resetDrawerData()
    router.refresh()
    return true
  }, [router])

  // Callbacks de Dependencias (criam/excluem direto, sem entrar no payload do submit)
  const dependenciasProps = useMemo(() => {
    if (!drawer || drawer.mode !== 'edit') return undefined
    if (drawer.tipo !== 'processo' && drawer.tipo !== 'atividade') return undefined

    const isProcesso = drawer.tipo === 'processo'
    const outrosNos: NoOption[] = initialNodes
      .filter((n) => n.tipo === drawer.tipo && n.id !== drawer.id)
      .map((n) => ({ id: n.id, descricao: n.descricao }))

    return {
      outrosNos,
      saidas: drawerDepsSaidas,
      entradas: drawerDepsEntradas,
      onAdd: async (input: { destinoId: number; tipo: DependenciaTipo; descricao?: string }) => {
        const action = isProcesso ? criarDependenciaProcesso : criarDependenciaAtividade
        const res = await action({
          origemId: drawer.id,
          destinoId: input.destinoId,
          tipo: input.tipo,
          descricao: input.descricao,
        })
        if (!res.success) return false
        const refresh = isProcesso ? listarDependenciasDeProcesso : listarDependenciasDeAtividade
        const next = await refresh(drawer.id)
        setDrawerDepsSaidas(next.saidas)
        setDrawerDepsEntradas(next.entradas)
        router.refresh()
        return true
      },
      onRemove: async (id: string) => {
        const action = isProcesso ? excluirDependenciaProcesso : excluirDependenciaAtividade
        const res = await action(id)
        if (!res.success) return false
        const refresh = isProcesso ? listarDependenciasDeProcesso : listarDependenciasDeAtividade
        const next = await refresh(drawer.id)
        setDrawerDepsSaidas(next.saidas)
        setDrawerDepsEntradas(next.entradas)
        router.refresh()
        return true
      },
    }
  }, [drawer, drawerDepsSaidas, drawerDepsEntradas, initialNodes, router])

  useEffect(() => {
    const map = dragTimers.current
    return () => map.forEach((t) => clearTimeout(t))
  }, [])

  return (
    <div className="relative w-full h-full">
      <ReactFlow
        nodes={nodes}
        edges={edges}
        nodeTypes={nodeTypes}
        onNodesChange={onNodesChange}
        onNodeClick={onNodeClick}
        fitView
        proOptions={{ hideAttribution: true }}
      >
        <Background color="#E2E8F0" gap={20} />
        <Controls position="bottom-right" showInteractive={false} />
        <MiniMap
          pannable zoomable
          nodeColor={() => '#1A6E8E'}
          maskColor="rgba(11,61,92,0.06)"
        />
      </ReactFlow>

      <button
        onClick={() => openDrawerFor({ mode: 'create', tipoFilho: 'cadeia' })}
        className="absolute top-4 left-4 bg-navy hover:bg-teal text-white text-sm font-semibold px-4 py-2 rounded-md shadow-md transition-all hover:-translate-y-0.5"
      >
        + Cadeia de Valor
      </button>

      {drawer && (
        <NodeDrawer
          state={drawer}
          transacoesDisponiveis={transacoes}
          initialTransacaoIds={drawerTransacoes}
          pessoasDisponiveis={pessoas}
          initialRaci={drawerRaci}
          produtosDisponiveis={produtos}
          initialProdutoIds={drawerProdutos}
          insumosDisponiveis={insumos}
          initialInsumos={drawerInsumos}
          sistemasDisponiveis={sistemas}
          initialSistemas={drawerSistemas}
          dependencias={dependenciasProps}
          onClose={() => {
            setDrawer(null)
            resetDrawerData()
          }}
          onSubmit={handleSubmit}
        />
      )}
    </div>
  )
}

function childTipo(tipo: NodeType): NodeType | null {
  const order: NodeType[] = ['cadeia', 'macroprocesso', 'processo', 'macroatividade', 'atividade']
  const idx = order.indexOf(tipo)
  return idx >= 0 && idx < order.length - 1 ? order[idx + 1] : null
}
