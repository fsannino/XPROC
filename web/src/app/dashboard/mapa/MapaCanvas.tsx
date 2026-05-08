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
import { MAPA_LEVELS, type NodeType } from '@/lib/definitions'
import NodeDrawer, { type DrawerState } from '@/components/mapa/NodeDrawer'
import CustomNode from '@/components/mapa/CustomNode'
import type { MultiSelectOption } from '@/components/ui/multi-select'

type Props = {
  initialNodes: MapaNode[]
  initialEdges: MapaEdge[]
  transacoes: MultiSelectOption[]
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

function CanvasInner({ initialNodes, initialEdges, transacoes }: Props) {
  const router = useRouter()

  // Mantém índice por id para o onNodeClick recuperar KPIs sem refetch
  const nodesById = useMemo(() => {
    const m = new Map<string, MapaNode>()
    for (const n of initialNodes) m.set(nodeKey(n.tipo, n.id), n)
    return m
  }, [initialNodes])

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

  const initialFlowEdges: Edge[] = useMemo(
    () =>
      initialEdges.map((e) => ({
        id: `e:${e.source}->${e.target}`,
        source: e.source,
        target: e.target,
        type: 'smoothstep',
        style: { stroke: '#1A6E8E', strokeWidth: 1.5 },
      })),
    [initialEdges],
  )

  const [nodes, setNodes] = useState<Node[]>(grouped)
  const [edges] = useState<Edge[]>(initialFlowEdges)
  const [drawer, setDrawer] = useState<DrawerState | null>(null)
  const [drawerTransacoes, setDrawerTransacoes] = useState<string[] | null>(null)
  const dragTimers = useRef(new Map<string, ReturnType<typeof setTimeout>>())

  const openDrawerFor = useCallback((s: DrawerState) => {
    setDrawer(s)
    setDrawerTransacoes(null)
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
    if (tipo === 'processo') {
      setDrawerTransacoes(null)
      try {
        const ids = await getTransacoesDeProcesso(id)
        setDrawerTransacoes(ids)
      } catch {
        setDrawerTransacoes([])
      }
    } else if (tipo === 'atividade') {
      setDrawerTransacoes(null)
      try {
        const ids = await getTransacoesDeAtividade(id)
        setDrawerTransacoes(ids)
      } catch {
        setDrawerTransacoes([])
      }
    } else {
      setDrawerTransacoes(null)
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

    if (payload.transacaoIds && payload.id && payload.mode === 'edit') {
      if (payload.tipo === 'processo') {
        await setProcessoTransacoes({ processoId: payload.id, transacaoIds: payload.transacaoIds })
      } else if (payload.tipo === 'atividade') {
        await setAtividadeTransacoes({ atividadeId: payload.id, transacaoIds: payload.transacaoIds })
      }
    }

    setDrawer(null)
    setDrawerTransacoes(null)
    router.refresh()
    return true
  }, [router])

  useEffect(() => {
    const map = dragTimers.current
    return () => map.forEach((t) => clearTimeout(t))
  }, [])

  void nodesById // (referenciado para evitar lint warning; reservado para futuras consultas)

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
          onClose={() => {
            setDrawer(null)
            setDrawerTransacoes(null)
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
